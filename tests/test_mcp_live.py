"""Live end-to-end test: MCP server to PowerShell bridge to Word to EndNote CWYW.

This test is skipped unless BOTH of these are set:

    ENDNOTE_LIVE_TEST=1
    ENDNOTE_TEST_LIBRARY=<absolute path to a real .enl library>

It drives probe, precheck, render, validate, set-style and validate again, and
cross-checks the rendered DOCX with the offline OOXML readers so a COM report that
disagrees with the file on disk is caught.

Nothing here writes to the library, and the input DOCX is built in a temporary
directory, so the only user-visible side effect is EndNote and Word being started.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TESTS_DIRECTORY = Path(__file__).resolve().parent
if str(TESTS_DIRECTORY) not in sys.path:
    sys.path.insert(0, str(TESTS_DIRECTORY))

from test_offline_structure import (  # noqa: E402  (path is set up above)
    count_citation_fields,
    count_bibliography_fields,
    paragraph_xml,
    read_layout_style,
    read_library_record_ids,
    text_run_xml,
    write_minimal_docx,
)

ROOT = TESTS_DIRECTORY.parent
SERVER = ROOT / "scripts" / "mcp_server.py"
LIVE_MANIFEST = TESTS_DIRECTORY / "fixtures" / "live-manifest.json"

LIVE_FLAG_VARIABLE = "ENDNOTE_LIVE_TEST"
LIBRARY_VARIABLE = "ENDNOTE_TEST_LIBRARY"
STYLE_VARIABLE = "ENDNOTE_TEST_STYLE"
TARGET_STYLE_VARIABLE = "ENDNOTE_TEST_TARGET_STYLE"

DEFAULT_TARGET_STYLE = "IEEE"
PROTOCOL_VERSION = "2025-06-18"
PROCESS_SHUTDOWN_TIMEOUT_SECONDS = 10


def _live_requested() -> bool:
    return os.environ.get(LIVE_FLAG_VARIABLE) == "1"


@unittest.skipUnless(_live_requested(), f"{LIVE_FLAG_VARIABLE} is not 1")
class McpLiveEndToEndTest(unittest.TestCase):
    """Full render, validate and restyle cycle against the installed EndNote."""

    def setUp(self) -> None:
        library = os.environ.get(LIBRARY_VARIABLE, "").strip()
        if not library:
            self.skipTest(f"{LIBRARY_VARIABLE} is not set")
        self.library = Path(library)
        if not self.library.is_file():
            self.skipTest(f"{LIBRARY_VARIABLE} does not point at a file: {self.library}")
        data_directory = self.library.with_suffix(".Data")
        if not data_directory.is_dir():
            self.skipTest(f"{data_directory} is missing, so the library cannot be read")

        self._directory = tempfile.TemporaryDirectory(prefix="endnote-live-")
        self.addCleanup(self._directory.cleanup)
        self.workspace = Path(self._directory.name)

        self.manifest = json.loads(LIVE_MANIFEST.read_text(encoding="utf-8"))
        self.manifest["library"]["path"] = str(self.library)
        override_style = os.environ.get(STYLE_VARIABLE, "").strip()
        if override_style:
            self.manifest["style"] = override_style
        self.manifest_path = self.workspace / "live-manifest.json"
        self.manifest_path.write_text(
            json.dumps(self.manifest, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        self.target_style = (
            os.environ.get(TARGET_STYLE_VARIABLE, "").strip() or DEFAULT_TARGET_STYLE
        )

        self.next_id = 0
        # The server relays the whole bridge log to its own stderr, and a live render logs
        # far more than the Windows pipe buffer holds. Nothing here reads that pipe, so it
        # is inherited rather than captured; otherwise the server blocks on a full buffer
        # and this test hangs forever. Inheriting also puts the bridge log where whoever
        # runs the live suite can read it.
        self.process = subprocess.Popen(
            [sys.executable, str(SERVER)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=None,
            text=True,
            encoding="utf-8",
        )
        self.addCleanup(self._shutdown)
        self._request(
            "initialize",
            {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "live-test", "version": "1"},
            },
        )

    def _shutdown(self) -> None:
        if self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=PROCESS_SHUTDOWN_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                self.process.kill()
        for stream in (self.process.stdin, self.process.stdout, self.process.stderr):
            if stream is not None:
                stream.close()

    def _request(self, method: str, params: dict) -> dict:
        assert self.process.stdin is not None
        assert self.process.stdout is not None
        self.next_id += 1
        payload = {"jsonrpc": "2.0", "id": self.next_id, "method": method, "params": params}
        self.process.stdin.write(json.dumps(payload, separators=(",", ":")) + "\n")
        self.process.stdin.flush()
        line = self.process.stdout.readline()
        self.assertTrue(line, f"MCP server closed stdout while handling {method}")
        return json.loads(line)

    def call(self, tool: str, arguments: dict) -> dict:
        response = self._request("tools/call", {"name": tool, "arguments": arguments})
        self.assertNotIn("error", response, f"{tool} returned a protocol error: {response}")
        return response["result"]["structuredContent"]

    def _build_input_document(self) -> Path:
        paragraphs = [
            paragraph_xml(
                text_run_xml("Live integration draft for the endnote-for-word bridge.")
            )
        ]
        for citation in self.manifest["citations"]:
            paragraphs.append(
                paragraph_xml(
                    text_run_xml(f"Citation {citation['id']} appears here "),
                    text_run_xml(citation["marker"]),
                    text_run_xml(" and the sentence ends."),
                )
            )
        return write_minimal_docx(self.workspace / "input.docx", "".join(paragraphs))

    def test_render_validate_and_restyle(self) -> None:
        probe = self.call("probe_endnote_word", {"library_path": str(self.library)})
        self.assertTrue(probe["ok"], probe)
        checks = probe["checks"]
        for name in (
            "windows",
            "word_com",
            "endnote_executable",
            "cwyw_addin_registered",
            "cwyw_template_present",
            "cwyw_template_loaded",
            "cwyw_macro_callable",
            "library_exists",
            "library_data_exists",
        ):
            self.assertTrue(checks[name], f"probe check '{name}' failed: {probe}")

        for style in (self.manifest["style"], self.target_style):
            styles = self.call("list_endnote_styles", {"style": style})
            self.assertTrue(styles["ok"], styles)
            installed = {name.casefold() for name in styles["styles"]}
            if style.casefold() not in installed:
                self.skipTest(f"output style '{style}' is not installed on this machine")

        precheck = self.call(
            "precheck_endnote_manifest",
            {"library_path": str(self.library), "manifest_path": str(self.manifest_path)},
        )
        if not precheck["ok"]:
            self.skipTest(
                "the fixture record numbers do not exist in "
                f"{self.library}: {precheck.get('problems')}"
            )

        input_path = self._build_input_document()
        original_input_bytes = input_path.read_bytes()
        output_path = self.workspace / "output.docx"
        render = self.call(
            "render_endnote_word_document",
            {
                "input_path": str(input_path),
                "output_path": str(output_path),
                "manifest_path": str(self.manifest_path),
            },
        )
        self.assertTrue(render["ok"], render)
        self.assertTrue(output_path.is_file(), "render reported success but wrote no file")
        self.assertEqual(
            input_path.read_bytes(),
            original_input_bytes,
            "the source document must not be modified by a render",
        )

        expected_citations = len(self.manifest["citations"])
        validation = self.call(
            "validate_endnote_word_document",
            {"document_path": str(output_path), "manifest_path": str(self.manifest_path)},
        )
        self.assertTrue(validation["ok"], validation)
        self.assertEqual(validation["errors"], [])
        self.assertEqual(validation["counts"]["native_citation_fields"], expected_citations)
        self.assertEqual(validation["counts"]["bibliography_fields"], 1)

        # Cross-check the COM report against the file on disk.
        self.assertEqual(count_citation_fields(output_path), expected_citations)
        self.assertEqual(count_bibliography_fields(output_path), 1)
        self.assertEqual(
            (read_layout_style(output_path) or "").casefold(),
            self.manifest["style"].casefold(),
        )
        expected_records = sorted(
            {
                target["rec_number"]
                for citation in self.manifest["citations"]
                for target in citation["targets"]
            }
        )
        self.assertEqual(read_library_record_ids(output_path), expected_records)

        restyled_path = self.workspace / "restyled.docx"
        restyle = self.call(
            "set_endnote_citation_style",
            {
                "input_path": str(output_path),
                "output_path": str(restyled_path),
                "style": self.target_style,
            },
        )
        self.assertTrue(restyle["ok"], restyle)
        # The bridge reports the resolved .ens base name, whose casing need not match the
        # name that was requested.
        self.assertEqual(restyle["style"].casefold(), self.target_style.casefold())
        self.assertEqual(
            restyle["previous_style"].casefold(), self.manifest["style"].casefold()
        )

        self.manifest["style"] = self.target_style
        self.manifest_path.write_text(
            json.dumps(self.manifest, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        revalidation = self.call(
            "validate_endnote_word_document",
            {"document_path": str(restyled_path), "manifest_path": str(self.manifest_path)},
        )
        self.assertTrue(revalidation["ok"], revalidation)
        self.assertEqual(revalidation["errors"], [])
        self.assertEqual(
            revalidation["counts"]["native_citation_fields"], expected_citations
        )
        self.assertEqual(
            (read_layout_style(restyled_path) or "").casefold(), self.target_style.casefold()
        )


if __name__ == "__main__":
    unittest.main()
