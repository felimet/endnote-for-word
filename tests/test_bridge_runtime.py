"""Offline runtime tests for bridge helpers that need a PowerShell host but not Word.

Invoke-HiddenProcess is lifted out of the bridge through the PowerShell AST and run against a
Python child. The UNC test drives import-ris and render into failures that happen before
EndNote or Word is looked up and that echo or read the resolved path. The library tests use
synthetic libraries and a per-test %APPDATA%, so the user's registered defaults are never
read or written.
"""

from __future__ import annotations

import ctypes
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TESTS_DIRECTORY = Path(__file__).resolve().parent
if str(TESTS_DIRECTORY) not in sys.path:
    sys.path.insert(0, str(TESTS_DIRECTORY))

from test_endnote_library import FIXTURE_RIS, build_synthetic_library  # noqa: E402
from test_offline_structure import (  # noqa: E402
    field_xml,
    paragraph_xml,
    text_run_xml,
    write_minimal_docx,
)

ROOT = TESTS_DIRECTORY.parent
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"
PYTHON_ENV_VARIABLE = "ENDNOTE_WORD_PYTHON"
NO_RIS_RECORDS_ERROR = "RisPath contains no RIS records (no line beginning with TY  - )."
NO_ENDNOTE_ERROR = (
    "EndNote is not installed or could not be discovered, so RIS cannot be imported through "
    "EndNote Direct Export."
)
ENDNOTE_INSTALL_GLOBS = ("EndNote*\\EndNote.exe",)

POWERSHELL_HOSTS = ("powershell.exe", "pwsh.exe")
CHILD_TIMEOUT_SECONDS = 60
# Far beyond the 4 KB anonymous pipe buffer that made sequential reads deadlock.
CHILD_STDERR_BYTES = 1024 * 1024

HIDDEN_PROCESS_HARNESS = r"""
param([Parameter(Mandatory)][string]$Bridge, [Parameter(Mandatory)][string]$Python, [Parameter(Mandatory)][int]$StderrBytes)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
function Write-Log { param([string]$Level, [string]$Message) [Console]::Error.WriteLine("[$Level] $Message") }
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
$definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-HiddenProcess' }, $true)
Invoke-Expression $definition.Extent.Text
$arguments = '-c "import sys; sys.stderr.write(''x'' * ' + $StderrBytes + '); sys.stderr.flush(); print(''child-done'')"'
$captured = Invoke-HiddenProcess -Executable $Python -Arguments $arguments
[ordered]@{ exit_code = $captured.exit_code; stdout = $captured.stdout.Trim(); stderr_length = $captured.stderr.Length } | ConvertTo-Json -Compress
"""


# Word hands a field's w:fldData back as raw bytes packed two per UTF-16 character, lone
# surrogates included; the harness packs a payload file the same way, without any encoder,
# and decodes it with the bridge's own helper.
FIELD_DATA_HARNESS = r"""
param([Parameter(Mandatory)][string]$Bridge, [Parameter(Mandatory)][string]$Payload)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
function Write-Log { param([string]$Level, [string]$Message) [Console]::Error.WriteLine("[$Level] $Message") }
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
$definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-FieldDataText' }, $true)
Invoke-Expression $definition.Extent.Text
$bytes = [IO.File]::ReadAllBytes($Payload)
$units = New-Object char[] ($bytes.Length / 2)
[Buffer]::BlockCopy($bytes, 0, $units, 0, $bytes.Length)
$packed = New-Object string (, $units)
[ordered]@{
    text  = Get-FieldDataText -Field ([pscustomobject]@{ Data = $packed }) -Label 'packed'
    empty = Get-FieldDataText -Field ([pscustomobject]@{ Data = '' }) -Label 'empty'
} | ConvertTo-Json -Compress
"""


# Find-UniqueMarkerRange against stand-ins for Word's Range and Find: one Find re-matches the
# same hit forever, the other matches a new hit every time. Both must end in a named error.
MARKER_GUARD_HARNESS = r"""
param([Parameter(Mandatory)][string]$Bridge)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
function Write-Log { param([string]$Level, [string]$Message) [Console]::Error.WriteLine("[$Level] $Message") }
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
foreach ($name in @('Release-ComObject', 'Find-UniqueMarkerRange')) {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Invoke-Expression $definition.Extent.Text
}
foreach ($variable in @('WdFindStop', 'WdCollapseEnd', 'MarkerSearchIterationLimit')) {
    $assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$script:' + $variable) }, $true)
    Invoke-Expression $assignment.Extent.Text
}
function New-FakeDocument {
    param([bool]$Advance)
    $range = [pscustomobject]@{ Start = 10; End = 20; Advance = $Advance }
    $find = [pscustomobject]@{ Text = ''; Forward = $true; Wrap = 0; MatchWildcards = $false; Range = $range }
    $find | Add-Member -MemberType ScriptMethod -Name ClearFormatting -Value { }
    $find | Add-Member -MemberType ScriptMethod -Name Execute -Value {
        if ($this.Range.Advance) { $this.Range.Start = $this.Range.End; $this.Range.End = $this.Range.Start + 10 }
        else { $this.Range.Start = 10; $this.Range.End = 20 }
        return $true
    }
    $range | Add-Member -MemberType NoteProperty -Name Find -Value $find
    $range | Add-Member -MemberType ScriptMethod -Name Collapse -Value { param($Direction) $this.Start = $this.End }
    return [pscustomobject]@{ Content = [pscustomobject]@{ Duplicate = $range } }
}
$result = [ordered]@{}
foreach ($case in @([pscustomobject]@{ name = 'stuck'; advance = $false }, [pscustomobject]@{ name = 'endless'; advance = $true })) {
    $script:Document = New-FakeDocument -Advance $case.advance
    try {
        $null = Find-UniqueMarkerRange -Marker ('{{CITE:' + $case.name + '}}')
        $result[$case.name] = 'returned'
    }
    catch { $result[$case.name] = $_.Exception.Message }
}
$result | ConvertTo-Json -Compress
"""


# Loads the named bridge functions and $script: constants, then runs the test body. The body
# reads its cases from a JSON file ($spec) and prints one JSON object.
FUNCTION_HARNESS = r"""
param([Parameter(Mandatory)][string]$Bridge, [Parameter(Mandatory)][string]$Cases)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
function Write-Log { param([string]$Level, [string]$Message) [Console]::Error.WriteLine("[$Level] $Message") }
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
$spec = [IO.File]::ReadAllText($Cases, [Text.Encoding]::UTF8) | ConvertFrom-Json
foreach ($name in $spec.functions) {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    Invoke-Expression $definition.Extent.Text
}
foreach ($variable in $spec.variables) {
    $assignment = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$script:' + $variable) }, $true)
    Invoke-Expression $assignment.Extent.Text
}
$result = [ordered]@{}
"""

MANIFEST_TARGET_BODY = r"""
foreach ($case in $spec.targets.PSObject.Properties) {
    try {
        Assert-ManifestTarget -Target $case.Value -CitationId $case.Name
        $result[$case.Name] = 'accepted'
    }
    catch { $result[$case.Name] = $_.Exception.Message }
}
foreach ($bibliography in $spec.bibliographies.PSObject.Properties) {
    $fields = [pscustomobject]@{
        bibliography_codes          = @('ADDIN EN.REFLIST')
        bibliography_result_lengths = @($bibliography.Value.Length)
        bibliography_results        = @($bibliography.Value)
    }
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    Add-BibliographyValidationProblems -Fields $fields -Manifest $spec.manifest -Errors $errors -Warnings $warnings
    $result[$bibliography.Name] = @($errors)
}
foreach ($bibliography in $spec.doiSuffixBibliographies.PSObject.Properties) {
    $fields = [pscustomobject]@{
        bibliography_codes          = @('ADDIN EN.REFLIST')
        bibliography_result_lengths = @($bibliography.Value.Length)
        bibliography_results        = @($bibliography.Value)
    }
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    Add-BibliographyValidationProblems -Fields $fields -Manifest $spec.doiSuffixManifest -Errors $errors -Warnings $warnings
    $result[$bibliography.Name] = @($errors)
}
foreach ($payloads in $spec.payloads.PSObject.Properties) {
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    Add-TravelingLibraryValidationProblems -Payloads @($payloads.Value) -Manifest $spec.manifest -Errors $errors -Warnings $warnings
    $result[$payloads.Name] = @($errors)
}
$result | ConvertTo-Json -Compress -Depth 5
"""

IMPORT_CONFIRMATION_BODY = r"""
foreach ($case in $spec.cases.PSObject.Properties) {
    $holders = $null
    if ($null -ne $case.Value.holders) { $holders = [int[]]@($case.Value.holders) }
    $titles = $null
    if ($null -ne $case.Value.titles) { $titles = [string[]]@($case.Value.titles) }
    $confirmation = Resolve-ImportTargetConfirmation -LibraryFile $spec.library -Windows @($case.Value.windows) `
        -RecentLibrary $case.Value.recent -HolderProcessIds $holders -LibraryWindowTitles $titles
    if ($null -eq $confirmation) { $result[$case.Name] = $null }
    else { $result[$case.Name] = '{0}:{1}' -f $confirmation.confirmed_by, $confirmation.process_id }
}
$result | ConvertTo-Json -Compress
"""

FILE_HOLDER_BODY = r"""
foreach ($path in @($spec.held, $spec.free)) {
    $ids = Get-FileHolderProcessIds -Path $path
    $result[[IO.Path]::GetFileName($path)] = $ids
}
# A hidden window of this process, so the window list is checked without EndNote.
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.Text = 'EndNote 21 - Probe Library.enl'
$null = $form.Handle
$result['own_windows'] = @([EndNoteForWordAgent.TopLevelWindows]::GetTitles([int[]]@($PID), 'WindowsForms'))
$result['other_class'] = @([EndNoteForWordAgent.TopLevelWindows]::GetTitles([int[]]@($PID), 'Afx:'))
# A same-class window that is no library, like EndNote 21.5's hidden com.endnote.async_tasks.
$tasks = New-Object System.Windows.Forms.Form
$tasks.Text = 'com.endnote.async_tasks'
$null = $tasks.Handle
function Get-Process { [CmdletBinding()] param([string]$Name) [pscustomobject]@{ Id = $PID } }
$script:LibraryWindowClassPrefix = 'WindowsForms'
$result['library_windows'] = Get-EndNoteLibraryWindowTitles
$tasks.Dispose()
$form.Dispose()
$result | ConvertTo-Json -Compress
"""


def available_hosts() -> list[str]:
    return [path for path in (shutil.which(name) for name in POWERSHELL_HOSTS) if path]


def run_host(host: str, arguments: list[str], environment: dict[str, str] | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(
        [host, "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", *arguments],
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=CHILD_TIMEOUT_SECONDS,
        check=False,
        # Never let a test start the real PowerShell 7 auto-install on a machine without pwsh.
        env={**os.environ, **(environment or {}), "ENDNOTE_WORD_AUTO_INSTALL_PWSH": "0"},
    )


def last_json_line(completed: subprocess.CompletedProcess) -> dict:
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        raise AssertionError(f"The bridge printed no JSON result. stderr:\n{completed.stderr}")
    return json.loads(lines[-1])


def short_path(path: Path) -> str | None:
    """Return the 8.3 alias of an existing path, or None when the volume generates none."""
    if sys.platform != "win32":
        return None
    buffer = ctypes.create_unicode_buffer(32768)
    if not ctypes.windll.kernel32.GetShortPathNameW(str(path), buffer, len(buffer)):
        return None
    return None if buffer.value.lower() == str(path).lower() else buffer.value


class BridgeRuntimeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-runtime-")
        cls.workspace = Path(cls._directory.name)
        cls.harness = cls.workspace / "hidden-process-harness.ps1"
        cls.harness.write_text(HIDDEN_PROCESS_HARNESS, encoding="utf-8-sig")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def test_hidden_process_survives_a_child_that_floods_stderr(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                try:
                    completed = run_host(
                        host,
                        [
                            "-File", str(self.harness),
                            "-Bridge", str(BRIDGE),
                            "-Python", sys.executable,
                            "-StderrBytes", str(CHILD_STDERR_BYTES),
                        ],
                    )
                except subprocess.TimeoutExpired:
                    self.fail("Invoke-HiddenProcess deadlocked on a child that floods stderr.")
                self.assertEqual(completed.returncode, 0, completed.stderr)
                result = last_json_line(completed)
                self.assertEqual(result["exit_code"], 0)
                self.assertEqual(result["stdout"], "child-done")
                self.assertEqual(result["stderr_length"], CHILD_STDERR_BYTES)

    def test_field_data_is_decoded_to_the_utf8_endnote_payload(self) -> None:
        # An Arabic surname starting at an odd byte offset packs into lone UTF-16 surrogates,
        # and the odd total length makes EndNote pad the payload with one NUL.
        surname = "عبدالله"
        xml = (
            f"<EndNote><Cite><Author>{surname}</Author><Year>2025</Year><RecNum>24</RecNum><record>"
            f"<rec-number>24</rec-number><author>{surname}, A.</author><author>Müller, H</author>"
            "</record></Cite></EndNote>"
        )
        data = xml.encode("utf-8")
        self.assertEqual(len(data) % 2, 1)
        encoded_surname = surname.encode("utf-8")
        offsets = [index for index in range(len(data)) if data.startswith(encoded_surname, index)]
        self.assertEqual([offset % 2 for offset in offsets], [1, 1])
        payload = self.workspace / "field-data.bin"
        payload.write_bytes(data + b"\x00")
        harness = self.workspace / "field-data-harness.ps1"
        harness.write_text(FIELD_DATA_HARNESS, encoding="utf-8-sig")
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                completed = run_host(host, ["-File", str(harness), "-Bridge", str(BRIDGE), "-Payload", str(payload)])
                self.assertEqual(completed.returncode, 0, completed.stderr)
                self.assertEqual(last_json_line(completed), {"text": xml, "empty": ""})

    def test_marker_search_fails_naming_the_marker_when_find_cannot_finish(self) -> None:
        harness = self.workspace / "marker-guard-harness.ps1"
        harness.write_text(MARKER_GUARD_HARNESS, encoding="utf-8-sig")
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                completed = run_host(host, ["-File", str(harness), "-Bridge", str(BRIDGE)])
                self.assertEqual(completed.returncode, 0, completed.stderr)
                result = last_json_line(completed)
                self.assertEqual(
                    result["stuck"],
                    "Word Find did not advance past character 20 while searching for the marker: {{CITE:stuck}}",
                )
                self.assertEqual(
                    result["endless"],
                    "Marker search stopped after 1000 matches without reaching the end of the document: {{CITE:endless}}",
                )

    def test_unc_paths_resolve_to_plain_file_system_paths(self) -> None:
        resolved = str(self.workspace.resolve())
        if len(resolved) < 3 or resolved[1:3] != ":\\":
            self.skipTest(f"The workspace is not on a drive letter: {resolved}")
        unc_workspace = "\\\\localhost\\" + resolved[0] + "$" + resolved[2:]
        # os.path.isdir turns access denied into False; Path.is_dir raises on it.
        if not os.path.isdir(unc_workspace):
            self.skipTest("The administrative share \\\\localhost\\<drive>$ is not reachable.")
        contents = {
            "references.txt": "TY  - JOUR\nER  - \n",
            "empty.ris": "no records here\n",
            "manifest.json": "{not json",
            "input.docx": "",
        }
        unc = {}
        for name, text in contents.items():
            (self.workspace / name).write_text(text, encoding="utf-8")
            unc[name] = unc_workspace + "\\" + name
        library = str(self.workspace / "Library.enl")
        cases = {
            # The resolved path is echoed back in plain form.
            "extension": (
                ["-Action", "import-ris", "-LibraryPath", library, "-RisPath", unc["references.txt"]],
                f"RisPath must use the .ris extension: {unc['references.txt']}",
            ),
            # Reaching the record count proves .NET opened the UNC file; the provider-qualified
            # form is rejected on 5.1 and read as a relative path on 7.
            "read": (
                ["-Action", "import-ris", "-LibraryPath", library, "-RisPath", unc["empty.ris"]],
                NO_RIS_RECORDS_ERROR,
            ),
            "render": (
                [
                    "-Action", "render",
                    "-InputPath", unc["input.docx"],
                    "-OutputPath", str(self.workspace / "output" / "output.docx"),
                    "-ManifestPath", unc["manifest.json"],
                ],
                f"ManifestPath is not valid JSON ({unc['manifest.json']}):",
            ),
        }
        for host in self.hosts:
            for case, (arguments, expected) in cases.items():
                with self.subTest(host=Path(host).name, case=case):
                    result = last_json_line(run_host(host, ["-File", str(BRIDGE), *arguments]))
                    self.assertFalse(result["ok"])
                    self.assertTrue(result["error"].startswith(expected), result["error"])
                    self.assertNotIn("FileSystem::", result["error"])

    def test_unformat_refuses_an_output_that_aliases_its_input_through_a_short_name(self) -> None:
        directory = self.workspace / "endnote long directory name"
        directory.mkdir(exist_ok=True)
        document = directory / "paper.docx"
        document.write_bytes(b"")
        alias = short_path(directory)
        if alias is None:
            self.skipTest("The volume does not generate 8.3 names.")
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                completed = run_host(
                    host,
                    [
                        "-File", str(BRIDGE),
                        "-Action", "unformat",
                        "-InputPath", os.path.join(alias, document.name),
                        "-OutputPath", str(document),
                        "-Overwrite",
                    ],
                )
                result = last_json_line(completed)
                self.assertFalse(result["ok"])
                self.assertTrue(result["error"].startswith("unformat is destructive"), result["error"])


class BridgeFunctionTests(unittest.TestCase):
    """Single bridge functions lifted out through the AST and driven with plain values."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-functions-")
        cls.workspace = Path(cls._directory.name)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def run_functions(self, name: str, body: str, spec: dict) -> dict:
        harness = self.workspace / f"{name}.ps1"
        harness.write_text(FUNCTION_HARNESS + body, encoding="utf-8-sig")
        cases = self.workspace / f"{name}.json"
        cases.write_text(json.dumps(spec), encoding="utf-8")
        results = {}
        for host in self.hosts:
            completed = run_host(host, ["-File", str(harness), "-Bridge", str(BRIDGE), "-Cases", str(cases)])
            self.assertEqual(completed.returncode, 0, completed.stderr)
            results[Path(host).name] = last_json_line(completed)
        return results

    def test_manifest_no_doi_opt_in_is_explicit_and_skips_only_the_visible_doi_gate(self) -> None:
        doi_target = {"rec_number": 3, "author": "Wang", "year": 2022, "doi": "10.0000/fixture.0103"}
        no_doi_target = {"rec_number": 1, "author": "Chen", "year": 2024, "no_doi": True}
        doi_record = ("<EndNote><Cite><Author>Wang</Author><record><rec-number>3</rec-number><contributors><author>Wang, H."
                      "</author></contributors><electronic-resource-num>10.0000/fixture.0103</electronic-resource-num>"
                      "</record></Cite></EndNote>")
        spec = {
            "functions": ["Get-OptionalProperty", "Assert-AllowedKeys", "ConvertTo-ManifestInteger",
                          "Assert-CitationText", "Test-NoDoiTarget", "Assert-ManifestTarget",
                          "Add-BibliographyValidationProblems", "Add-TravelingLibraryValidationProblems"],
            "variables": ["AllowedTargetKeys", "NoDoiTargetKey", "DoiPattern", "ForbiddenCitationTextPattern",
                          "DoiResolverPrefix", "DuplicateDoiResolverPattern", "DanglingDoiResolverPattern",
                          "RecordNumberElementFormat", "TravelingRecordDoiPattern",
                          "TravelingRecordDoiElsewherePattern"],
            "targets": {
                "withDoi": doi_target,
                "noDoi": no_doi_target,
                "noDoiAndDoi": {**no_doi_target, "doi": "10.0000/fixture.0103"},
                "noDoiText": {**no_doi_target, "no_doi": "yes"},
                "noDoiFalse": {**no_doi_target, "no_doi": False},
                "missingDoi": {"rec_number": 1, "author": "Chen", "year": 2024},
            },
            "manifest": {
                "bibliography": {"expected": True},
                "citations": [{"id": "a", "targets": [doi_target]}, {"id": "b", "targets": [no_doi_target]}],
            },
            # A DOI target whose own suffix holds 'doi:'; DoiPattern accepts it.
            "doiSuffixManifest": {
                "bibliography": {"expected": True},
                "citations": [{"id": "c", "targets": [{**doi_target, "doi": "10.1000/abc-doi:5"}]}],
            },
            "bibliographies": {
                "shown": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r[2] M. Chen, Web page, 2024.",
                "hidden": "[1] H. Wang, Journal, 2022.\r[2] M. Chen, Web page, 2024.",
                "dangling": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                            "[2] M. Chen, Web page, 2024, doi: https://doi.org/.",
                # EndNote's template rules drop 'https://doi.org/' with the empty DOI field but keep
                # the space-separated 'doi:' label.
                "orphanLabel": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                               "[2] M. Chen, Web page. doi: https://example.org/page (accessed Nov. 1, 2026).",
                "orphanLabelAtEnd": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                                    "[2] M. Chen, Web page, 2024, doi: .",
                "orphanLabelMidList": "[1] M. Chen, Web page, 2024, doi:\r"
                                      "[2] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.",
                # The leftover label is followed by whatever the template prints next.
                "orphanBeforeOnline": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                                      "[2] M. Chen, Web page, 2024, doi: . [Online]. Available: https://example.org/page",
                "orphanUpperBeforeOnline": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                                           "[2] M. Chen, Web page, 2024, DOI: [Online]. Available: https://example.org/page",
                "orphanBeforeAccessed": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                                        "[2] M. Chen, Web page, 2024, doi: (accessed Nov. 1, 2026).",
                # An unnumbered style: the next entry starts with a letter.
                "orphanBeforeAuthorEntry": "Chen, M. (2024). Web page. doi:\r"
                                           "Wang, H. (2022). Journal. doi: https://doi.org/10.0000/fixture.0103.",
                # 'doi:' inside a token or before ordinary words is no leftover label.
                "doiInsideTokens": "[1] H. Wang, Journal, doi: https://doi.org/10.0000/fixture.0103.\r"
                                   "[2] M. Chen, Why doi: names persist, Web page, urn:doi:10.1000/abc, 2024.",
            },
            "doiSuffixBibliographies": {
                "doiSuffix": "[1] H. Wang, Journal, doi: https://doi.org/10.1000/abc-doi:5.",
            },
            "payloads": {
                # Shapes measured in an EndNote 21 document: a DOI record carries
                # <electronic-resource-num>, a web page without a DOI carries none.
                "noDoiRecordClean": [
                    "<EndNote><Cite><Author>Wang</Author><record><rec-number>3</rec-number><contributors><author>Wang, H.</author>"
                    "</contributors><electronic-resource-num>10.0000/fixture.0103</electronic-resource-num></record></Cite>"
                    "<Cite><Author>Chen</Author><record><rec-number>1</rec-number><ref-type name=\"Web Page\">12</ref-type>"
                    "<contributors><author>Chen, M.</author></contributors><electronic-resource-num></electronic-resource-num>"
                    "<language>eng</language></record></Cite></EndNote>",
                ],
                # The no_doi record comes first in one merged payload: only its own record is read.
                "noDoiRecordBeforeDoiRecord": [
                    "<EndNote><Cite><Author>Chen</Author><record><rec-number>1</rec-number><ref-type name=\"Web Page\">12</ref-type>"
                    "<contributors><author>Chen, M.</author></contributors><urls><related-urls><url>https://example.org/page</url>"
                    "</related-urls></urls></record></Cite><Cite><Author>Wang</Author><record><rec-number>3</rec-number>"
                    "<contributors><author>Wang, H.</author></contributors><urls><related-urls><url>https://doi.org/10.0000/fixture.0103"
                    "</url></related-urls></urls><electronic-resource-num>10.0000/fixture.0103</electronic-resource-num>"
                    "</record></Cite></EndNote>",
                ],
                "noDoiRecordDoiInUrl": [
                    doi_record,
                    "<EndNote><Cite><Author>Chen</Author><record><rec-number>1</rec-number><contributors><author>Chen, M.</author>"
                    "</contributors><urls><related-urls><url><style face=\"normal\">https://doi.org/10.0000/fixture.0101</style>"
                    "</url></related-urls></urls></record></Cite></EndNote>",
                ],
                "noDoiRecordDoiInNotes": [
                    doi_record,
                    "<EndNote><Cite><Author>Chen</Author><record><rec-number>1</rec-number><contributors><author>Chen, M.</author>"
                    "</contributors><notes>Imported. doi:10.0000/fixture.0101</notes></record></Cite></EndNote>",
                ],
                "noDoiRecordHasDoi": [
                    "<EndNote><Cite><Author>Wang</Author><record><rec-number>3</rec-number><contributors><author>Wang, H.</author>"
                    "</contributors><electronic-resource-num>10.0000/fixture.0103</electronic-resource-num></record></Cite></EndNote>",
                    "<EndNote><Cite><Author>Chen</Author><record><rec-number>1</rec-number><contributors><author>Chen, M.</author>"
                    "</contributors><electronic-resource-num><style face=\"normal\">10.0000/fixture.0101</style>"
                    "</electronic-resource-num></record></Cite></EndNote>",
                ],
            },
        }
        missing = "Citation '{0}' requires a DOI for every target, or \"no_doi\": true for a record that has none."
        dangling = ("The bibliography prints a 'doi:' label without 'https://doi.org/<DOI>' after it: '{0}'. Tie the "
                    "output style's 'doi:' label to the DOI field with Link Adjacent Text so it disappears for a "
                    "record without a DOI.")
        elsewhere = ("Citation 'b': rec-number 1 is marked no_doi, but its traveling library record carries the DOI "
                     "'10.0000/fixture.0101' in <{0}>. Move it into the record's DOI field in EndNote, remove no_doi "
                     "and cite the target with doi instead.")
        expected = {
            "withDoi": "accepted",
            "noDoi": "accepted",
            "noDoiAndDoi": "Citation 'noDoiAndDoi' sets no_doi and also carries doi; keep only the DOI.",
            "noDoiText": "Citation 'noDoiText' no_doi must be true or false.",
            "noDoiFalse": missing.format("noDoiFalse"),
            "missingDoi": missing.format("missingDoi"),
            "shown": [],
            "hidden": [
                "The bibliography does not visibly contain 'doi: https://doi.org/10.0000/fixture.0103'. "
                "Check the output style Bibliography Templates."
            ],
            "dangling": [dangling.format("doi: https://doi.org/.")],
            "orphanLabel": [dangling.format("doi: https://example.org/page (accessed ")],
            "orphanLabelAtEnd": [dangling.format("doi: .")],
            "orphanLabelMidList": [dangling.format("doi:\r[2] H. Wang, Journal, doi: https://")],
            "orphanBeforeOnline": [dangling.format("doi: . [Online]. Available: https://exam")],
            "orphanUpperBeforeOnline": [dangling.format("DOI: [Online]. Available: https://exampl")],
            "orphanBeforeAccessed": [dangling.format("doi: (accessed Nov. 1, 2026).")],
            "orphanBeforeAuthorEntry": [dangling.format("doi:\rWang, H. (2022). Journal. doi: http")],
            "doiInsideTokens": [],
            "doiSuffix": [],
            "noDoiRecordClean": [],
            "noDoiRecordBeforeDoiRecord": [],
            "noDoiRecordDoiInUrl": [elsewhere.format("url")],
            "noDoiRecordDoiInNotes": [elsewhere.format("notes")],
            "noDoiRecordHasDoi": [
                "Citation 'b': rec-number 1 is marked no_doi, but its traveling library record carries the DOI "
                "'10.0000/fixture.0101'. Remove no_doi, cite that DOI and make the style print it."
            ],
        }
        for host, result in self.run_functions("manifest-target", MANIFEST_TARGET_BODY, spec).items():
            with self.subTest(host=host):
                self.assertEqual(result, expected)

    def test_import_target_needs_its_title_and_a_full_path_signal_from_the_same_process(self) -> None:
        library = r"E:\Projects\Grading\Lib [2024].enl"
        front = [{"process_id": 42, "title": "EndNote 21 - Lib [2024].enl"}]
        other = r"E:\Other\My EndNote Library.enl"
        # Library windows behind the front one: the start-up library and the target.
        open_libraries = ["EndNote 21 - Lib [2024].enl", "EndNote 21 - My EndNote Library.enl"]
        spec = {
            "functions": ["Resolve-ImportTargetConfirmation"],
            "variables": ["EndNoteProcessName"],
            "library": library,
            "cases": {
                # Measured failure of 1.5.2: the window is up, File1 still names the start-up library.
                "handleWhileRecentNamesAnother": {"windows": front, "recent": other, "holders": [42], "titles": open_libraries},
                "titleWithoutProductPrefix": {"windows": [{"process_id": 42, "title": "lib [2024].ENL"}], "recent": None,
                                              "holders": [42], "titles": ["lib [2024].ENL"]},
                "recentLibraryWithoutRestartManager": {"windows": front, "recent": library.upper(), "holders": None, "titles": None},
                "heldByAnotherProcessOnly": {"windows": front, "recent": other, "holders": [7], "titles": open_libraries},
                "heldButAnotherLibraryInFront": {"windows": [{"process_id": 42, "title": "EndNote 21 - Other.enl"}], "recent": library,
                                                 "holders": [42], "titles": open_libraries},
                "sameSuffixOtherName": {"windows": [{"process_id": 42, "title": "EndNote 21 - MyLib [2024].enl"}], "recent": library,
                                        "holders": [42], "titles": ["EndNote 21 - MyLib [2024].enl"]},
                "noEndNoteWindow": {"windows": [], "recent": library, "holders": [42], "titles": []},
                # A same-name library from another folder is open too: the one in front may be it.
                "sameNameLibraryAlsoOpen": {"windows": front, "recent": other, "holders": [42],
                                            "titles": ["EndNote 21 - Lib [2024].enl", "Lib [2024].enl"]},
                # Restart Manager answered: the target is not held, so File1 must not confirm it.
                "recentLibraryButNotHeld": {"windows": front, "recent": library, "holders": [], "titles": ["EndNote 21 - Lib [2024].enl"]},
                "recentLibraryHeldElsewhere": {"windows": front, "recent": library, "holders": [7], "titles": ["EndNote 21 - Lib [2024].enl"]},
                # No window list: the same-name check cannot run, so the handle does not confirm.
                "handleWithoutWindowList": {"windows": front, "recent": other, "holders": [42], "titles": None},
                # The front title is not among the library windows found: the list is not trusted.
                "handleWithEmptyWindowList": {"windows": front, "recent": other, "holders": [42], "titles": []},
                # 'Project - Lib [2024].enl' is another file, so it neither blocks nor passes for the target.
                "longerNameEndingInTarget": {"windows": front, "recent": other, "holders": [42],
                                             "titles": ["EndNote 21 - Lib [2024].enl", "EndNote 21 - Project - Lib [2024].enl"]},
                "longerNameInFront": {"windows": [{"process_id": 42, "title": "EndNote 21 - Project - Lib [2024].enl"}],
                                      "recent": library, "holders": [42], "titles": ["EndNote 21 - Project - Lib [2024].enl"]},
            },
        }
        expected = {
            "handleWhileRecentNamesAnother": "open_handle:42",
            "titleWithoutProductPrefix": "open_handle:42",
            "recentLibraryWithoutRestartManager": "recent_library:42",
            "heldByAnotherProcessOnly": None,
            "heldButAnotherLibraryInFront": None,
            "sameSuffixOtherName": None,
            "noEndNoteWindow": None,
            "sameNameLibraryAlsoOpen": None,
            "recentLibraryButNotHeld": None,
            "recentLibraryHeldElsewhere": None,
            "handleWithoutWindowList": None,
            "handleWithEmptyWindowList": None,
            "longerNameEndingInTarget": "open_handle:42",
            "longerNameInFront": None,
        }
        for host, result in self.run_functions("import-confirmation", IMPORT_CONFIRMATION_BODY, spec).items():
            with self.subTest(host=host):
                self.assertEqual(result, expected)

    def test_native_helpers_name_file_holders_and_list_windows_by_class(self) -> None:
        held = self.workspace / "held.txt"
        free = self.workspace / "free.txt"
        held.write_text("held", encoding="utf-8")
        free.write_text("free", encoding="utf-8")
        spec = {
            "functions": ["Test-NativeHelpers", "Get-FileHolderProcessIds", "Get-EndNoteLibraryWindowTitles"],
            "variables": ["RestartManagerSource", "RestartManagerTypeName", "RestartManagerUnavailable", "RestartManagerWarning",
                          "EndNoteProcessName", "LibraryExtension"],
            "held": str(held),
            "free": str(free),
        }
        with held.open("rb"):
            results = self.run_functions("file-holders", FILE_HOLDER_BODY, spec)
        for host, result in results.items():
            with self.subTest(host=host):
                self.assertIn(os.getpid(), result["held.txt"])
                self.assertEqual(result["free.txt"], [])
                self.assertIn("EndNote 21 - Probe Library.enl", result["own_windows"])
                self.assertEqual(result["other_class"], [])
                self.assertEqual(result["library_windows"], ["EndNote 21 - Probe Library.enl"])


class DocumentLibraryTests(unittest.TestCase):
    """configure-defaults template registration and import-ris without LibraryPath."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-doclib-")
        cls.workspace = Path(cls._directory.name)
        cls.empty_template = build_synthetic_library(cls.workspace / "template", "Empty", records=())
        cls.populated_library = build_synthetic_library(cls.workspace / "populated", "Populated")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def run_bridge(self, host: str, name: str, *arguments: str) -> dict:
        # A fresh %APPDATA% per host and test isolates the registered defaults.
        appdata = self.workspace / "appdata" / f"{name}-{Path(host).stem}"
        appdata.mkdir(parents=True, exist_ok=True)
        environment = {"APPDATA": str(appdata), PYTHON_ENV_VARIABLE: sys.executable}
        return last_json_line(run_host(host, ["-File", str(BRIDGE), *arguments], environment))

    def plain_document(self, name: str) -> Path:
        return write_minimal_docx(self.workspace / name, paragraph_xml(text_run_xml("No citations yet.")))

    def test_configure_defaults_registers_and_keeps_an_empty_library_template(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                registered = self.run_bridge(
                    host, "keep", "-Action", "configure-defaults",
                    "-EmptyLibraryTemplatePath", str(self.empty_template),
                )
                self.assertTrue(registered["ok"], registered)
                self.assertEqual(registered["defaults"]["library"]["empty_template_path"], str(self.empty_template))
                kept = self.run_bridge(host, "keep", "-Action", "configure-defaults")
                self.assertTrue(kept["ok"], kept)
                self.assertEqual(kept["defaults"]["library"]["empty_template_path"], str(self.empty_template))

    def test_configure_defaults_rejects_a_template_that_holds_references(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                result = self.run_bridge(
                    host, "reject", "-Action", "configure-defaults",
                    "-EmptyLibraryTemplatePath", str(self.populated_library),
                )
                self.assertFalse(result["ok"])
                self.assertTrue(
                    result["error"].startswith("The empty library template must contain zero stored references"),
                    result["error"],
                )

    def test_import_without_a_library_targets_the_library_beside_the_document(self) -> None:
        document = self.plain_document("draft paper.docx")
        expected = self.workspace / "draft paper.enl"
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                result = self.run_bridge(
                    host, "derive", "-Action", "import-ris",
                    "-InputPath", str(document), "-RisPath", str(FIXTURE_RIS),
                )
                self.assertFalse(result["ok"])
                self.assertTrue(
                    result["error"].startswith(f"No EndNote library exists at {expected}."), result["error"]
                )

    def test_import_refuses_to_create_a_library_for_a_document_that_already_cites(self) -> None:
        # The split code is only recognisable once its w:instrText runs are joined.
        citation = field_xml(" ADDIN EN.CITE <EndNote><Cite><RecNum>1</RecNum></Cite></EndNote> ", "[1]", split_code=True)
        document = write_minimal_docx(self.workspace / "cited.docx", paragraph_xml(citation))
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                result = self.run_bridge(
                    host, "refuse", "-Action", "import-ris",
                    "-InputPath", str(document), "-RisPath", str(FIXTURE_RIS),
                )
                self.assertFalse(result["ok"])
                self.assertTrue(result["error"].startswith("InputPath already carries EndNote fields"), result["error"])
                self.assertFalse((self.workspace / "cited.enl").exists())

    def test_import_creates_the_document_library_from_the_registered_template(self) -> None:
        # Past the copy, import-ris opens EndNote, so this runs only on a CI runner without it.
        program_files = [os.environ.get(name) for name in ("ProgramFiles", "ProgramFiles(x86)")]
        installed = any(
            glob.glob(os.path.join(root, pattern)) for root in program_files if root for pattern in ENDNOTE_INSTALL_GLOBS
        )
        if os.environ.get("GITHUB_ACTIONS") != "true" or installed:
            self.skipTest("Needs a machine without EndNote, so the import stops right after the library copy.")
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                document = self.plain_document(f"new paper {Path(host).stem}.docx")
                library = document.with_suffix(".enl")
                registered = self.run_bridge(
                    host, "create", "-Action", "configure-defaults",
                    "-EmptyLibraryTemplatePath", str(self.empty_template),
                )
                self.assertTrue(registered["ok"], registered)
                result = self.run_bridge(
                    host, "create", "-Action", "import-ris",
                    "-InputPath", str(document), "-RisPath", str(FIXTURE_RIS),
                )
                self.assertEqual(result["error"], NO_ENDNOTE_ERROR)
                self.assertTrue(library.is_file())
                self.assertTrue((library.with_suffix(".Data") / "sdb" / "sdb.eni").is_file())


if __name__ == "__main__":
    unittest.main(verbosity=2)
