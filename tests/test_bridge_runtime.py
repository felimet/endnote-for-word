"""Offline runtime tests for bridge helpers that need a PowerShell host but not Word.

Invoke-HiddenProcess is lifted out of the bridge through the PowerShell AST and run against a
Python child. The UNC test drives -Action import-ris, which rejects a non-.ris file before
EndNote is ever looked up, and echoes the resolved path in its error. The library tests use
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

    def test_unc_paths_resolve_to_plain_file_system_paths(self) -> None:
        local = self.workspace / "references.txt"
        local.write_text("TY  - JOUR\nER  - \n", encoding="utf-8")
        resolved = str(local.resolve())
        if len(resolved) < 3 or resolved[1:3] != ":\\":
            self.skipTest(f"The workspace is not on a drive letter: {resolved}")
        unc = "\\\\localhost\\" + resolved[0] + "$" + resolved[2:]
        # os.path.isfile turns access denied into False; Path.is_file raises on it.
        if not os.path.isfile(unc):
            self.skipTest("The administrative share \\\\localhost\\<drive>$ is not reachable.")
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                completed = run_host(
                    host,
                    [
                        "-File", str(BRIDGE),
                        "-Action", "import-ris",
                        "-LibraryPath", str(self.workspace / "Library.enl"),
                        "-RisPath", unc,
                    ],
                )
                result = last_json_line(completed)
                self.assertFalse(result["ok"])
                self.assertEqual(result["error"], f"RisPath must use the .ris extension: {unc}")

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
