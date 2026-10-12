"""Offline runtime tests for bridge helpers that need a PowerShell host but not Word.

Invoke-HiddenProcess is lifted out of the bridge through the PowerShell AST and run against a
Python child.
"""

from __future__ import annotations

import json
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
    )


def last_json_line(completed: subprocess.CompletedProcess) -> dict:
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        raise AssertionError(f"The bridge printed no JSON result. stderr:\n{completed.stderr}")
    return json.loads(lines[-1])



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


if __name__ == "__main__":
    unittest.main(verbosity=2)
