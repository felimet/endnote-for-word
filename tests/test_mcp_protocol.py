"""Protocol-level tests for scripts/mcp_server.py.

The server is driven over stdio exactly as an MCP client would drive it. The one test
that reaches the PowerShell bridge points ENDNOTE_WORD_POWERSHELL at a stub executable
that ignores its arguments and echoes a fixed JSON line, so no real PowerShell, Word or
EndNote process is ever started.
"""

from __future__ import annotations

import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SERVER = ROOT / "scripts" / "mcp_server.py"
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"

POWERSHELL_ENV_VARIABLE = "ENDNOTE_WORD_POWERSHELL"
PROTOCOL_VERSION = "2025-06-18"
SERVER_NAME = "endnote-for-word"
SERVER_VERSION = "1.5.1"
PROCESS_SHUTDOWN_TIMEOUT_SECONDS = 5

STUB_RESULT = {"ok": True, "action": "stub", "checks": {"stub_bridge": True}}

EXPECTED_TOOLS = {
    "probe_endnote_word",
    "list_endnote_styles",
    "search_endnote_library",
    "get_endnote_reference",
    "precheck_endnote_manifest",
    "configure_endnote_defaults",
    "render_endnote_word_document",
    "validate_endnote_word_document",
    "update_endnote_word_fields",
    "set_endnote_citation_style",
    "unformat_endnote_word_document",
    "export_reference_import_file",
    "import_reference_file",
    "get_endnote_job_status",
    "cancel_endnote_job",
}

DESTRUCTIVE_TOOLS = {
    "render_endnote_word_document",
    "update_endnote_word_fields",
    "set_endnote_citation_style",
    "unformat_endnote_word_document",
    "import_reference_file",
}

# Build spec section 0.6b: ENRemoveFieldCodes raises a modal confirmation dialog and hangs
# under automation, so the tool it would back must never reappear in the catalogue.
FORBIDDEN_TOOLS = {"remove_endnote_field_codes"}

READ_ONLY_TOOLS = {
    "probe_endnote_word",
    "list_endnote_styles",
    "search_endnote_library",
    "get_endnote_reference",
    "precheck_endnote_manifest",
    "validate_endnote_word_document",
    "get_endnote_job_status",
}

REQUIRED_ANNOTATIONS = ("readOnlyHint", "destructiveHint", "idempotentHint", "openWorldHint")


def load_server_module():
    """Load the server module without starting its stdio loop."""
    spec = importlib.util.spec_from_file_location("endnote_for_word_mcp_server", SERVER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load {SERVER}")
    module = importlib.util.module_from_spec(spec)
    # dataclasses resolves postponed annotations through sys.modules, so register first.
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def write_stub_powershell(directory: Path) -> Path:
    """Write an executable that ignores every argument and echoes one JSON line."""
    payload = json.dumps(STUB_RESULT, separators=(", ", ": "))
    if os.name == "nt":
        stub = directory / "stub-powershell.cmd"
        stub.write_text(f"@echo off\r\necho {payload}\r\n", encoding="ascii")
    else:
        stub = directory / "stub-powershell.sh"
        stub.write_text(f"#!/bin/sh\necho '{payload}'\n", encoding="ascii")
        stub.chmod(0o755)
    return stub


def terminate(process: subprocess.Popen) -> None:
    """Shut a spawned server down without leaking pipes."""
    if process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=PROCESS_SHUTDOWN_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            process.kill()
    for stream in (process.stdin, process.stdout, process.stderr):
        if stream is not None:
            stream.close()


class McpProtocolTests(unittest.TestCase):
    """Handshake, tool catalogue and error mapping."""

    @classmethod
    def setUpClass(cls) -> None:
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-mcp-")
        cls.stub = write_stub_powershell(Path(cls._directory.name))

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def setUp(self) -> None:
        environment = os.environ.copy()
        environment[POWERSHELL_ENV_VARIABLE] = str(self.stub)
        environment["PYTHONIOENCODING"] = "utf-8"
        self.process = subprocess.Popen(
            [sys.executable, str(SERVER)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=environment,
            text=True,
            encoding="utf-8",
        )
        self.addCleanup(terminate, self.process)

    def request(self, payload: dict) -> dict:
        assert self.process.stdin is not None
        assert self.process.stdout is not None
        self.process.stdin.write(json.dumps(payload, separators=(",", ":")) + "\n")
        self.process.stdin.flush()
        line = self.process.stdout.readline()
        self.assertTrue(line, "MCP server closed stdout")
        return json.loads(line)

    def initialize(self) -> dict:
        return self.request(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": PROTOCOL_VERSION,
                    "capabilities": {},
                    "clientInfo": {"name": "unit-test", "version": "1"},
                },
            }
        )

    def list_tools(self) -> list[dict]:
        self.initialize()
        response = self.request({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
        return response["result"]["tools"]

    def test_initialize_reports_the_server_identity(self) -> None:
        result = self.initialize()["result"]
        self.assertEqual(result["serverInfo"]["name"], SERVER_NAME)
        self.assertEqual(result["serverInfo"]["version"], SERVER_VERSION)
        self.assertIn("tools", result["capabilities"])
        self.assertEqual(result["protocolVersion"], PROTOCOL_VERSION)

    def test_tool_catalogue_matches_the_contract(self) -> None:
        names = {tool["name"] for tool in self.list_tools()}
        self.assertEqual(names, EXPECTED_TOOLS)
        self.assertEqual(names & FORBIDDEN_TOOLS, set(), "a tool dropped after measurement is back")

    def test_every_tool_schema_is_closed_and_annotated(self) -> None:
        for tool in self.list_tools():
            with self.subTest(tool=tool["name"]):
                self.assertTrue(tool["description"].strip())
                schema = tool["inputSchema"]
                self.assertEqual(schema["type"], "object")
                self.assertIs(schema["additionalProperties"], False)
                self.assertIsInstance(schema["required"], list)
                self.assertIsInstance(schema.get("properties", {}), dict)
                for name in schema["required"]:
                    self.assertIn(name, schema["properties"])
                annotations = tool["annotations"]
                for hint in REQUIRED_ANNOTATIONS:
                    self.assertIn(hint, annotations)
                    self.assertIsInstance(annotations[hint], bool)

    def test_destructive_and_read_only_hints_are_correct(self) -> None:
        annotations = {tool["name"]: tool["annotations"] for tool in self.list_tools()}
        for name in DESTRUCTIVE_TOOLS:
            with self.subTest(tool=name):
                self.assertTrue(annotations[name]["destructiveHint"])
                self.assertFalse(annotations[name]["readOnlyHint"])
        for name in READ_ONLY_TOOLS:
            with self.subTest(tool=name):
                self.assertTrue(annotations[name]["readOnlyHint"])
                self.assertFalse(annotations[name]["destructiveHint"])

        self.assertTrue(annotations["precheck_endnote_manifest"]["openWorldHint"])

    def test_ris_export_requires_a_doi(self) -> None:
        tools = {tool["name"]: tool for tool in self.list_tools()}
        item_schema = tools["export_reference_import_file"]["inputSchema"]["properties"][
            "references"
        ]["items"]
        self.assertIn("doi", item_schema["required"])

    def test_ris_import_targets_an_explicit_library_and_file(self) -> None:
        tools = {tool["name"]: tool for tool in self.list_tools()}
        schema = tools["import_reference_file"]["inputSchema"]
        # library_path may be omitted in favour of document_path's library.
        self.assertEqual(schema["required"], ["ris_path"])
        self.assertIn("library_path", schema["properties"])
        self.assertIn("document_path", schema["properties"])
        self.assertIn("empty_library_template_path", schema["properties"])
        configure = tools["configure_endnote_defaults"]["inputSchema"]
        self.assertIn("empty_library_template_path", configure["properties"])
        annotations = tools["import_reference_file"]["annotations"]
        self.assertTrue(annotations["destructiveHint"])
        self.assertFalse(annotations["idempotentHint"])

    def test_ris_import_arguments_reach_the_powershell_bridge(self) -> None:
        server = load_server_module()
        command = server._build_command(
            "import-ris",
            {
                "library_path": r"C:\Research\Target.enl",
                "ris_path": r"C:\Research\references.ris",
                "empty_library_template_path": r"C:\Research\Empty.enl",
            },
            "pwsh.exe",
        )
        self.assertEqual(command[0], "pwsh.exe")
        self.assertEqual(command[command.index("-LibraryPath") + 1], r"C:\Research\Target.enl")
        self.assertEqual(command[command.index("-RisPath") + 1], r"C:\Research\references.ris")
        self.assertEqual(
            command[command.index("-EmptyLibraryTemplatePath") + 1],
            r"C:\Research\Empty.enl",
        )
        derived = server._build_command(
            "import-ris",
            {"document_path": r"C:\Research\paper.docx", "ris_path": r"C:\Research\references.ris"},
            "pwsh.exe",
        )
        self.assertNotIn("-LibraryPath", derived)
        self.assertEqual(derived[derived.index("-InputPath") + 1], r"C:\Research\paper.docx")

    def test_word_document_markers_name_every_document_the_bridge_opened(self) -> None:
        server = load_server_module()
        stderr = (
            "[INFO] WORD_PID=4242\r\n"
            "[INFO] WORD_DOCUMENT=D:\\論文\\.草稿.0123.work.docx\r\n"
            "[INFO] Saving temporary DOCX: D:\\論文\\.草稿.4567.tmp.docx\r\n"
            "[INFO] WORD_DOCUMENT=D:\\論文\\.草稿.4567.tmp.docx\r\n"
            "[WARN] text mentioning WORD_DOCUMENT=C:\\not\\a\\marker.docx\r\n"
        )
        self.assertEqual(
            server._extract_word_documents(stderr),
            {"D:\\論文\\.草稿.0123.work.docx", "D:\\論文\\.草稿.4567.tmp.docx"},
        )
        # Without a marker nothing is removed and no recovery state is read.
        self.assertEqual(server._clear_killed_word_recovery(4242, "[INFO] WORD_PID=4242\r\n", set()), [])

    def test_default_render_timeout_scales_for_large_manifests(self) -> None:
        server = load_server_module()
        previous = os.environ.pop("ENDNOTE_WORD_TIMEOUT_SECONDS", None)
        try:
            # The bridge and the precheck accept an optional BOM, so the budget must as well.
            for encoding in ("utf-8", "utf-8-sig"):
                with self.subTest(encoding=encoding):
                    with tempfile.TemporaryDirectory(prefix="endnote-timeout-") as directory:
                        manifest = Path(directory) / "manifest.json"
                        manifest.write_text(
                            json.dumps({"citations": [{"id": str(index)} for index in range(83)]}),
                            encoding=encoding,
                        )
                        timeout = server._timeout_seconds("render", {"manifest_path": str(manifest)})
                    self.assertEqual(timeout, 724)
        finally:
            if previous is not None:
                os.environ["ENDNOTE_WORD_TIMEOUT_SECONDS"] = previous

    def test_powershell_notice_survives_a_bridge_timeout(self) -> None:
        if os.name != "nt":
            self.skipTest("The stub bridge is a Windows batch file.")
        server = load_server_module()
        notice = {"status": "installing", "message": "PowerShell 7 is being installed."}
        previous = os.environ.get("ENDNOTE_WORD_POWERSHELL")
        try:
            with tempfile.TemporaryDirectory(prefix="endnote-notice-") as directory:
                stub = Path(directory, "slow-bridge.cmd")
                stub.write_text(
                    "@echo off\r\n"
                    f"echo POWERSHELL_NOTICE={json.dumps(notice)} 1>&2\r\n"
                    "ping -n 4 127.0.0.1 >nul\r\n",
                    encoding="ascii",
                )
                os.environ["ENDNOTE_WORD_POWERSHELL"] = str(stub)
                result = server._invoke_bridge("probe", {}, timeout=1)
        finally:
            if previous is None:
                os.environ.pop("ENDNOTE_WORD_POWERSHELL", None)
            else:
                os.environ["ENDNOTE_WORD_POWERSHELL"] = previous
        self.assertFalse(result["ok"])
        self.assertIn("timed out", result["error"])
        self.assertEqual(result["powershell_notice"], notice)

    def test_powershell_7_is_found_at_its_known_location_without_path(self) -> None:
        server = load_server_module()
        names = ("ENDNOTE_WORD_POWERSHELL", "PATH", "LOCALAPPDATA")
        previous = {name: os.environ.get(name) for name in names}
        try:
            with tempfile.TemporaryDirectory(prefix="endnote-pwsh7-") as directory:
                alias = Path(directory, "Microsoft", "WindowsApps", "pwsh.exe")
                alias.parent.mkdir(parents=True)
                alias.write_bytes(b"")
                os.environ.pop("ENDNOTE_WORD_POWERSHELL", None)
                os.environ["PATH"] = str(Path(directory, "empty"))
                os.environ["LOCALAPPDATA"] = directory
                self.assertEqual(server._powershell_executable(), str(alias))
        finally:
            for name, value in previous.items():
                if value is None:
                    os.environ.pop(name, None)
                else:
                    os.environ[name] = value

    def test_explicit_timeout_overrides_manifest_scaling(self) -> None:
        server = load_server_module()
        previous = os.environ.get("ENDNOTE_WORD_TIMEOUT_SECONDS")
        os.environ["ENDNOTE_WORD_TIMEOUT_SECONDS"] = "300"
        try:
            timeout = server._timeout_seconds("render", {"manifest_path": "does-not-matter.json"})
        finally:
            if previous is None:
                os.environ.pop("ENDNOTE_WORD_TIMEOUT_SECONDS", None)
            else:
                os.environ["ENDNOTE_WORD_TIMEOUT_SECONDS"] = previous
        self.assertEqual(timeout, 300)

    def test_ping_and_unknown_method(self) -> None:
        self.assertEqual(self.request({"jsonrpc": "2.0", "id": 3, "method": "ping"})["result"], {})
        missing = self.request({"jsonrpc": "2.0", "id": 4, "method": "does/not/exist"})
        self.assertEqual(missing["error"]["code"], -32601)

    def test_unknown_tool_is_reported_as_an_error(self) -> None:
        self.initialize()
        response = self.request(
            {
                "jsonrpc": "2.0",
                "id": 5,
                "method": "tools/call",
                "params": {"name": "does_not_exist", "arguments": {}},
            }
        )
        self.assertEqual(response["id"], 5)
        self.assertEqual(response["error"]["code"], -32603)

    def test_bridge_tool_uses_the_configured_powershell_executable(self) -> None:
        if not BRIDGE.is_file():
            self.skipTest(f"{BRIDGE} is missing, so the bridge cannot be invoked.")
        self.initialize()
        response = self.request(
            {
                "jsonrpc": "2.0",
                "id": 6,
                "method": "tools/call",
                "params": {"name": "probe_endnote_word", "arguments": {}},
            }
        )
        result = response["result"]
        self.assertFalse(result["isError"], result)
        self.assertEqual(result["structuredContent"], STUB_RESULT)

    def call_tool(self, request_id: int, name: str, arguments: dict) -> dict:
        response = self.request(
            {"jsonrpc": "2.0", "id": request_id, "method": "tools/call", "params": {"name": name, "arguments": arguments}}
        )
        return response["result"]["structuredContent"]

    def render_arguments(self) -> dict:
        # The manifest names no existing library, so the worker opens no EndNote session.
        return {
            "input_path": str(Path(self._directory.name) / "input.docx"),
            "output_path": str(Path(self._directory.name) / "output.docx"),
            "manifest_path": str(Path(self._directory.name) / "missing-manifest.json"),
        }

    def test_render_returns_a_job_that_finishes_with_the_bridge_result(self) -> None:
        self.initialize()
        submitted = self.call_tool(10, "render_endnote_word_document", self.render_arguments())
        self.assertTrue(submitted["ok"], submitted)
        self.assertIn(submitted["state"], {"queued", "waiting_for_lock", "running"})
        status = self.call_tool(11, "get_endnote_job_status", {"job_id": submitted["job_id"], "wait_seconds": 60})
        self.assertEqual(status["job"]["state"], "succeeded", status)
        self.assertEqual(status["job"]["result"], STUB_RESULT)
        listing = self.call_tool(12, "get_endnote_job_status", {})
        self.assertEqual([job["job_id"] for job in listing["jobs"]], [submitted["job_id"]])

    def test_background_false_waits_for_the_result(self) -> None:
        self.initialize()
        result = self.call_tool(20, "render_endnote_word_document", {**self.render_arguments(), "background": False})
        self.assertEqual(result, STUB_RESULT)

    def test_unknown_job_is_reported(self) -> None:
        self.initialize()
        status = self.call_tool(30, "get_endnote_job_status", {"job_id": "missing"})
        self.assertFalse(status["ok"])


@unittest.skipUnless(os.name == "nt", "The cancellation path ends the bridge with taskkill.")
class McpJobCancellationTests(unittest.TestCase):
    """A slow stub bridge keeps the first job running while the second waits in the queue."""

    def setUp(self) -> None:
        self._directory = tempfile.TemporaryDirectory(prefix="endnote-jobs-")
        self.addCleanup(self._directory.cleanup)
        payload = json.dumps(STUB_RESULT, separators=(", ", ": "))
        stub = Path(self._directory.name) / "slow-stub-powershell.cmd"
        stub.write_text(f"@echo off\r\nping -n 30 127.0.0.1 >nul\r\necho {payload}\r\n", encoding="ascii")
        environment = {**os.environ, POWERSHELL_ENV_VARIABLE: str(stub), "PYTHONIOENCODING": "utf-8"}
        self.process = subprocess.Popen(
            [sys.executable, str(SERVER)], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL, env=environment, text=True, encoding="utf-8",
        )
        self.addCleanup(terminate, self.process)

    def call_tool(self, request_id: int, name: str, arguments: dict) -> dict:
        message = {"jsonrpc": "2.0", "id": request_id, "method": "tools/call", "params": {"name": name, "arguments": arguments}}
        self.process.stdin.write(json.dumps(message) + "\n")
        self.process.stdin.flush()
        return json.loads(self.process.stdout.readline())["result"]["structuredContent"]

    def test_queued_and_running_jobs_can_be_cancelled(self) -> None:
        arguments = {
            "input_path": str(Path(self._directory.name) / "input.docx"),
            "output_path": str(Path(self._directory.name) / "output.docx"),
            "manifest_path": str(Path(self._directory.name) / "missing-manifest.json"),
        }
        running = self.call_tool(1, "render_endnote_word_document", arguments)["job_id"]
        queued = self.call_tool(2, "update_endnote_word_fields", {"input_path": arguments["input_path"]})["job_id"]
        cancelled = self.call_tool(3, "cancel_endnote_job", {"job_id": queued})
        self.assertEqual(cancelled["job"]["state"], "cancelled")
        self.call_tool(4, "cancel_endnote_job", {"job_id": running})
        status = self.call_tool(5, "get_endnote_job_status", {"job_id": running, "wait_seconds": 20})
        self.assertEqual(status["job"]["state"], "cancelled", status)
        self.assertFalse(status["job"]["result"]["ok"])


class OwnedEndNoteCloseTests(unittest.TestCase):
    """The idle close of an EndNote this server started, with the window layer faked."""

    class FakeSessions:
        POLL_SECONDS = 0

        def __init__(self, library_handles: list[int], frame_visible: bool) -> None:
            self.closed: list[int] = []
            self.restored: list[dict] = []
            self.library_handles = library_handles
            self.frame_visible = frame_visible

        def library_windows(self, _process_ids):
            return [types.SimpleNamespace(handle=handle, title=f"Library {handle}.enl") for handle in self.library_handles]

        def frame_windows(self, _process_ids):
            return [types.SimpleNamespace(handle=2, title="EndNote 21", visible=self.frame_visible)]

        def dismiss_notices(self, _process_ids):
            return []

        def restore_recent_libraries(self, snapshot: dict) -> None:
            self.restored.append(snapshot)

    class FreeLock:
        def acquire(self, _should_stop) -> bool:
            return True

        def release(self) -> None:
            pass

    def close_owned(self, exits_on_close: bool, library_handles: list[int] = (1,), frame_visible: bool = False):
        from unittest import mock

        server = load_server_module()
        sessions = self.FakeSessions(list(library_handles), frame_visible)
        running = {4242}
        killed: list[int] = []

        def close_window(handle: int) -> None:
            sessions.closed.append(handle)
            if exits_on_close:
                running.clear()

        def terminate(process_id: int) -> bool:
            killed.append(process_id)
            running.discard(process_id)
            return True

        sessions.close_window = close_window
        runner = server._JobRunner()
        runner._owned_endnote = (4242, {"File1": "mine.enl"})
        runner._owned_windows = {1}
        with mock.patch.multiple(server, LIBRARY_CLOSE_WAIT_SECONDS=0, ENDNOTE_EXIT_WAIT_SECONDS=0,
                                 _session_module=lambda: sessions, _terminate_process_pid=terminate,
                                 _process_pids=lambda _name: set(running)):
            runner._close_owned_endnote(self.FreeLock())
        return killed, sessions, runner

    def test_an_endnote_that_closes_normally_is_not_terminated(self) -> None:
        killed, sessions, runner = self.close_owned(exits_on_close=True)
        self.assertEqual(killed, [])
        self.assertEqual(sessions.closed, [1])
        self.assertIsNone(runner._owned_endnote)
        self.assertEqual(sessions.restored, [{"File1": "mine.enl"}])

    def test_an_owned_endnote_that_ignores_every_close_request_is_terminated(self) -> None:
        killed, sessions, runner = self.close_owned(exits_on_close=False)
        self.assertEqual(sessions.closed, [1, 2])
        self.assertEqual(killed, [4242])
        self.assertIsNone(runner._owned_endnote)
        self.assertEqual(sessions.restored, [{"File1": "mine.enl"}])

    def test_a_library_the_user_opened_in_it_makes_the_endnote_theirs(self) -> None:
        for library_handles, frame_visible in (([1, 7], False), ([1], True)):
            killed, sessions, runner = self.close_owned(False, library_handles, frame_visible)
            self.assertEqual((killed, sessions.closed, sessions.restored), ([], [], []))
            self.assertIsNone(runner._owned_endnote)


@unittest.skipUnless(os.name == "nt", "Only Windows refuses to rename a process's working folder.")
class ServerWorkingFolderTests(unittest.TestCase):
    """Codex starts the server in the installed plugin folder, which an update must replace."""

    def rename_while_serving(self, folder: Path, server: Path) -> None:
        process = subprocess.Popen(
            [sys.executable, str(server)], cwd=folder, stdin=subprocess.PIPE,
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, encoding="utf-8",
        )
        try:
            process.stdin.write('{"jsonrpc":"2.0","id":1,"method":"ping"}\n')
            process.stdin.flush()
            self.assertEqual(json.loads(process.stdout.readline())["id"], 1)
            folder.rename(folder.with_name(folder.name + "-renamed"))
        finally:
            terminate(process)

    def test_server_leaves_the_plugin_folder_it_was_started_in(self) -> None:
        with tempfile.TemporaryDirectory() as scratch:
            plugin = Path(scratch) / "plugin"
            (plugin / "scripts").mkdir(parents=True)
            shutil.copy2(SERVER, plugin / "scripts" / SERVER.name)
            self.rename_while_serving(plugin, plugin / "scripts" / SERVER.name)

    def test_server_keeps_any_other_folder_as_the_base_for_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as scratch:
            project = Path(scratch) / "project"
            project.mkdir()
            with self.assertRaises(PermissionError):
                self.rename_while_serving(project, SERVER)


class PluginLauncherTests(unittest.TestCase):
    """The .mcp.json launcher must resolve the plugin root from either variable."""

    def test_launcher_starts_from_either_plugin_root_variable(self) -> None:
        config = json.loads((ROOT / ".mcp.json").read_text(encoding="utf-8"))
        server = config["mcpServers"][SERVER_NAME]
        command = server["command"]
        if shutil.which(command) is None:
            self.skipTest(f"'{command}' from .mcp.json is not on PATH.")
        for variable in ("CLAUDE_PLUGIN_ROOT", "PLUGIN_ROOT"):
            with self.subTest(variable=variable):
                environment = os.environ.copy()
                environment.pop("CLAUDE_PLUGIN_ROOT", None)
                environment.pop("PLUGIN_ROOT", None)
                environment[variable] = str(ROOT)
                # Pin the interpreter so this test exercises plugin-root resolution only and
                # not whatever python shim happens to sit first on the machine's PATH.
                environment["ENDNOTE_WORD_PYTHON"] = sys.executable
                process = subprocess.Popen(
                    [command, *server["args"]],
                    cwd=ROOT.parent,
                    env=environment,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                )
                try:
                    assert process.stdin is not None
                    assert process.stdout is not None
                    process.stdin.write(
                        json.dumps({"jsonrpc": "2.0", "id": 9, "method": "ping"}) + "\n"
                    )
                    process.stdin.flush()
                    line = process.stdout.readline()
                    self.assertTrue(line, "launcher produced no response")
                    response = json.loads(line)
                    self.assertEqual(response["id"], 9)
                    self.assertEqual(response["result"], {})
                finally:
                    terminate(process)


if __name__ == "__main__":
    unittest.main()
