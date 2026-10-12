"""Offline runtime tests for bridge helpers that need a PowerShell host but not Word.

Invoke-HiddenProcess is lifted out of the bridge through the PowerShell AST and run against a
Python child. The UNC test drives -Action import-ris, which rejects a non-.ris file before
EndNote is ever looked up, and echoes the resolved path in its error.
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"

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


def run_host(host: str, arguments: list[str]) -> subprocess.CompletedProcess:
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
        env={**os.environ, "ENDNOTE_WORD_AUTO_INSTALL_PWSH": "0"},
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
