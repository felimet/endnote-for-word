"""Word-backed bridge tests: they need Word COM, but neither EndNote nor a library.

Skipped unless ENDNOTE_LIVE_TEST=1 (``tests\\run-tests.ps1 -IncludeLive``) and Word is
registered, so CI and the default offline run never start Word. Every document is a
synthetic DOCX opened read-only by a hidden Word of the test's own.
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

from test_bridge_runtime import BRIDGE, available_hosts, last_json_line, run_host  # noqa: E402
from test_offline_structure import (  # noqa: E402
    LAYOUT_VARIABLE_NAME,
    LIBRARIES_VARIABLE_NAME,
    _SAMPLE_LAYOUT,
    field_xml,
    paragraph_xml,
    table_xml,
    text_run_xml,
    write_minimal_docx,
)

LIVE_FLAG_VARIABLE = "ENDNOTE_LIVE_TEST"
WORD_PROGID = "Word.Application"

# Lifts Find-UniqueMarkerRange and the constants it reads out of the bridge through the AST,
# then looks every marker up in one Word session. WORD_PID lets the test end a Word that a
# hanging search left behind.
FIND_MARKER_HARNESS = r"""
param([Parameter(Mandatory)][string]$Bridge, [Parameter(Mandatory)][string]$DocumentPath, [Parameter(Mandatory)][string]$MarkerList)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
foreach ($name in @('Write-Log', 'Release-ComObject', 'Find-UniqueMarkerRange')) {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Invoke-Expression $definition.Extent.Text
}
foreach ($variable in @('WdFindStop', 'WdCollapseEnd', 'MarkerSearchIterationLimit')) {
    $assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$script:' + $variable) }, $true)
    Invoke-Expression $assignment.Extent.Text
}
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = New-Object -ComObject Word.Application
# Named only when unambiguous, like the bridge's Start-Word, so a timeout never ends another Word.
$started = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
if ($started.Count -eq 1) { [Console]::Error.WriteLine("WORD_PID=$($started[0].Id)") }
$results = [ordered]@{}
$script:Document = $null
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $script:Document = $word.Documents.Open($DocumentPath, $false, $true)
    foreach ($marker in $MarkerList.Split('|')) {
        try {
            $range = Find-UniqueMarkerRange -Marker $marker
            # 12 is WdInformation.wdWithInTable.
            $results[$marker] = [ordered]@{ text = [string]$range.Text; in_table = [bool]$range.Information(12) }
            Release-ComObject $range
        }
        catch { $results[$marker] = [ordered]@{ error = $_.Exception.Message } }
    }
}
finally {
    try { if ($null -ne $script:Document) { $script:Document.Close(0) } }
    finally { $word.Quit() }
}
$results | ConvertTo-Json -Compress
"""

# Real EndNote payload shape: UTF-8 XML with a non-ASCII author, NUL-padded to an even length.
DATA_PAYLOAD = (
    '<EndNote><Cite AuthorYear="1"><Author>Maier-Hein</Author><Year>2024</Year><RecNum>24</RecNum>'
    "<DisplayText>Maier-Hein, et al. [1]</DisplayText><record><rec-number>24</rec-number>"
    "<contributors><authors><author>Maier-Hein, Lena</author><author>Müller, Henning</author>"
    "</authors></contributors><titles><title>Metrics reloaded</title></titles></record></Cite></EndNote>"
)
# An Arabic surname whose UTF-8 bytes start at odd offsets packs into lone UTF-16 surrogates.
ARABIC_SURNAME = "عبدالله"
ARABIC_PAYLOAD = (
    f"<EndNote><Cite><Author>{ARABIC_SURNAME}</Author><Year>2025</Year><RecNum>25</RecNum>"
    "<DisplayText>[2]</DisplayText><record><rec-number>25</rec-number><volume>17</volume>"
    f"<contributors><authors><author>{ARABIC_SURNAME}, A.</author></authors></contributors></record></Cite></EndNote>"
)
DATA_LIBRARIES = (
    '<Libraries><item db-id="fixturedbid">FixtureLibrary<record-ids><item>24</item><item>25</item>'
    "</record-ids></item></Libraries>"
)


def word_is_registered() -> bool:
    if sys.platform != "win32":
        return False
    import winreg

    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, WORD_PROGID))
    except OSError:
        return False
    return True


def padded_payload(text: str) -> bytes:
    data = text.encode("utf-8")
    return data + b"\x00" * (len(data) % 2)


@unittest.skipUnless(os.environ.get(LIVE_FLAG_VARIABLE) == "1", f"{LIVE_FLAG_VARIABLE} is not 1")
class WordBridgeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if not word_is_registered():
            raise unittest.SkipTest(f"{WORD_PROGID} is not registered on this machine.")
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-word-")
        cls.workspace = Path(cls._directory.name)
        cls.harness = cls.workspace / "find-marker-harness.ps1"
        cls.harness.write_text(FIND_MARKER_HARNESS, encoding="utf-8-sig")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def run_word_host(self, host: str, arguments: list[str], stalled: str) -> subprocess.CompletedProcess:
        """Run a host that starts Word; on a timeout, end the Word it named and fail."""
        try:
            return run_host(host, arguments)
        except subprocess.TimeoutExpired as error:
            stderr = error.stderr.decode("utf-8", "replace") if isinstance(error.stderr, bytes) else error.stderr or ""
            for line in stderr.splitlines():
                if "WORD_PID=" in line:
                    subprocess.run(["taskkill", "/F", "/PID", line.split("WORD_PID=", 1)[1].strip()], capture_output=True, check=False)
            self.fail(stalled)

    def run_find_harness(self, host: str, document: Path, markers: list[str]) -> dict:
        arguments = ["-File", str(self.harness), "-Bridge", str(BRIDGE), "-DocumentPath", str(document)]
        completed = self.run_word_host(
            host, [*arguments, "-MarkerList", "|".join(markers)], "Find-UniqueMarkerRange did not return; a marker search is looping."
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        return last_json_line(completed)

    def test_marker_search_finishes_inside_table_cells(self) -> None:
        figure = table_xml(
            paragraph_xml(text_run_xml("[figure]")),
            paragraph_xml(text_run_xml("Fig. 1. Caption {{CITE:middle}} continues {{CITE:twice}} {{CITE:caption}}")),
        )
        document = write_minimal_docx(
            self.workspace / "table-markers.docx",
            paragraph_xml(text_run_xml("Before the figure."))
            + figure
            + paragraph_xml(text_run_xml("After {{CITE:body}} and again {{CITE:twice}}.")),
        )
        markers = ["{{CITE:caption}}", "{{CITE:middle}}", "{{CITE:body}}", "{{CITE:twice}}", "{{CITE:absent}}"]
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                found = self.run_find_harness(host, document, markers)
                for marker, in_table in (("{{CITE:caption}}", True), ("{{CITE:middle}}", True), ("{{CITE:body}}", False)):
                    self.assertEqual(found[marker], {"text": marker, "in_table": in_table})
                self.assertIn("found 2", found["{{CITE:twice}}"]["error"])
                self.assertIn("found 0", found["{{CITE:absent}}"]["error"])

    def validate(self, host: str, document: Path, author: str) -> dict:
        manifest = {
            "schema_version": 1,
            "style": "IEEE",
            "citations": [
                {
                    "id": "metrics",
                    "marker": "{{CITE:metrics}}",
                    "targets": [
                        {"rec_number": 24, "author": author, "year": 2024, "doi": "10.1038/s41592-023-02151-z"}
                    ],
                    "display": "author-year",
                },
                {
                    "id": "arabic",
                    "marker": "{{CITE:arabic}}",
                    "targets": [{"rec_number": 25, "author": ARABIC_SURNAME, "year": 2025, "doi": "10.1145/3698767"}],
                    "display": "default",
                },
            ],
        }
        manifest_path = self.workspace / f"manifest-{author}.json"
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        arguments = ["-File", str(BRIDGE), "-Action", "validate", "-InputPath", str(document), "-ManifestPath", str(manifest_path)]
        return last_json_line(self.run_word_host(host, arguments, "validate did not return."))

    def citation_with_field_data(self, payload: str, result: str) -> str:
        data = padded_payload(payload)
        nested = field_xml(" ADDIN EN.CITE.DATA ", data=data)
        return field_xml(" ADDIN EN.CITE ", result=result, nested=nested, data=data)

    def test_validate_reads_the_payload_endnote_keeps_in_field_data(self) -> None:
        arabic = ARABIC_PAYLOAD.encode("utf-8")
        encoded_surname = ARABIC_SURNAME.encode("utf-8")
        offsets = [index for index in range(len(arabic)) if arabic.startswith(encoded_surname, index)]
        self.assertEqual([offset % 2 for offset in offsets], [1, 1])
        document = write_minimal_docx(
            self.workspace / "field-data.docx",
            table_xml(
                paragraph_xml(text_run_xml("[figure]")),
                paragraph_xml(
                    text_run_xml("Fig. 1. As "),
                    self.citation_with_field_data(DATA_PAYLOAD, "Maier-Hein, et al. [1]"),
                    text_run_xml(" recommend "),
                    self.citation_with_field_data(ARABIC_PAYLOAD, "[2]"),
                    text_run_xml("."),
                ),
            )
            + paragraph_xml(text_run_xml("After the figure.")),
            {LAYOUT_VARIABLE_NAME: _SAMPLE_LAYOUT, LIBRARIES_VARIABLE_NAME: DATA_LIBRARIES},
        )
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                report = self.validate(host, document, "Maier-Hein")
                self.assertTrue(report["ok"], report)
                self.assertEqual(report["errors"], [])
                self.assertEqual(report["counts"]["native_citation_fields"], 2)
                self.assertEqual(report["counts"]["citation_data_fields"], 2)

                # The decoded record is really compared, not merely found.
                wrong = self.validate(host, document, "Rahmath")
                self.assertFalse(wrong["ok"])
                self.assertIn(
                    "Citation 'metrics': rec-number 24 was found, but the author surname 'Rahmath' does not appear in the same traveling library payload.",
                    wrong["errors"],
                )


if __name__ == "__main__":
    unittest.main()
