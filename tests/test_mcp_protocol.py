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
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SERVER = ROOT / "scripts" / "mcp_server.py"
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"

POWERSHELL_ENV_VARIABLE = "ENDNOTE_WORD_POWERSHELL"
PROTOCOL_VERSION = "2025-06-18"
SERVER_NAME = "endnote-for-word"
SERVER_VERSION = "1.2.0"
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
}

REQUIRED_ANNOTATIONS = ("readOnlyHint", "destructiveHint", "idempotentHint", "openWorldHint")


def load_server_module():
    """Load the server module without starting its stdio loop."""
    spec = importlib.util.spec_from_file_location("endnote_for_word_mcp_server", SERVER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load {SERVER}")
    module = importlib.util.module_from_spec(spec)
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
        self.assertEqual(schema["required"], ["library_path", "ris_path"])
        self.assertIn("empty_library_template_path", schema["properties"])
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

    def test_default_render_timeout_scales_for_large_manifests(self) -> None:
        server = load_server_module()
        previous = os.environ.pop("ENDNOTE_WORD_TIMEOUT_SECONDS", None)
        try:
            with tempfile.TemporaryDirectory(prefix="endnote-timeout-") as directory:
                manifest = Path(directory) / "manifest.json"
                manifest.write_text(
                    json.dumps({"citations": [{"id": str(index)} for index in range(83)]}),
                    encoding="utf-8",
                )
                timeout = server._timeout_seconds("render", {"manifest_path": str(manifest)})
        finally:
            if previous is not None:
                os.environ["ENDNOTE_WORD_TIMEOUT_SECONDS"] = previous
        self.assertEqual(timeout, 724)

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
