"""Offline tests for the bridge's PowerShell 7 auto-install state machine.

The functions are lifted out of scripts/endnote-word.ps1 through the PowerShell AST. Get-Command
is shadowed so each case controls whether pwsh.exe and winget.exe exist, the state directory
points into a temporary folder, and winget is a stub batch file that records its arguments and
writes the log. Nothing is installed and the real winget never runs. One end-to-end case runs
the real bridge with every PowerShell 7 and winget location hidden.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"
POWERSHELL_HOSTS = ("powershell.exe", "pwsh.exe")
HOST_TIMEOUT_SECONDS = 60
STUB_WAIT_SECONDS = 30
STUB_POLL_SECONDS = 0.2
# The slow stub runs for about eight seconds; a caller that waits for it would take longer.
CALLER_BUDGET_SECONDS = 5
EXPECTED_WINGET_ARGUMENTS = (
    "install --id Microsoft.PowerShell --exact --source winget --scope user --installer-type msix "
    "--no-upgrade --silent --disable-interactivity --accept-package-agreements "
    "--accept-source-agreements --log"
)
CONTROLLED_VARIABLES = ("ENDNOTE_WORD_AUTO_INSTALL_PWSH", "ENDNOTE_WORD_POWERSHELL")

HARNESS = r"""
param(
    [Parameter(Mandatory)][string]$Bridge,
    [Parameter(Mandatory)][string]$StateDirectory,
    [string]$Action = 'render',
    [string]$Winget = '',
    [string]$Pwsh = '',
    [string]$Edition = 'Desktop',
    [switch]$WingetStillRunning
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
function Write-Log { param([string]$Level, [string]$Message) [Console]::Error.WriteLine("[$Level] $Message") }
function Get-Command {
    param([string]$Name, $CommandType, $ErrorAction)
    if ($Name -eq 'winget.exe' -and $Winget) { return [pscustomobject]@{ Source = $Winget } }
    if ($Name -eq 'pwsh.exe' -and $Pwsh) { return [pscustomobject]@{ Source = $Pwsh } }
    return $null
}
if ($WingetStillRunning) {
    function Get-Process { param($Id, $ErrorAction) return [pscustomobject]@{ ProcessName = 'winget' } }
}
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Bridge, [ref]$null, [ref]$null)
$wanted = @('Get-OptionalProperty', 'ConvertTo-ProcessArgument', 'Read-StrictUtf8Text', 'Read-Utf8JsonFile',
    'Find-PowerShell7', 'Find-Winget', 'Read-Pwsh7InstallState', 'Write-Pwsh7InstallState', 'New-Pwsh7Notice',
    'Get-Pwsh7InstallLogTail', 'Complete-Pwsh7InstallState', 'Start-Pwsh7Install', 'Update-PowerShell7AutoInstall')
foreach ($definition in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($wanted -contains $definition.Name) { Invoke-Expression $definition.Extent.Text }
}
foreach ($statement in $ast.EndBlock.Statements) {
    if ($statement.Extent.Text -match '^\$script:(Pwsh7|WingetExecutableName|FalseSettingValues|PowerShellInstallUrl|PowerShellHostEnvironmentVariable|ReadOnlyActions)') {
        Invoke-Expression $statement.Extent.Text
    }
}
$script:Pwsh7StateDirectory = $StateDirectory
$script:Pwsh7InstallStatePath = Join-Path $StateDirectory 'pwsh-install.json'
$script:Pwsh7InstallLogPath = Join-Path $StateDirectory 'pwsh-install.log'
$script:Pwsh7KnownPaths = @()
$script:WingetKnownPaths = @()
[ordered]@{ notice = (Update-PowerShell7AutoInstall -Action $Action -Edition $Edition) } | ConvertTo-Json -Depth 5 -Compress
"""

WINGET_STUB = (
    "@echo off\r\n"
    "echo %*> \"%~dp0winget-arguments.txt\"\r\n"
    "if exist \"%~dp0slow\" ping -n 9 127.0.0.1 >nul\r\n"
    "echo stub winget finished> \"%~dp0pwsh-install.log\"\r\n"
    "echo done> \"%~dp0winget-done.txt\"\r\n"
    "exit /b 0\r\n"
)


def available_hosts() -> list[str]:
    return [path for path in (shutil.which(name) for name in POWERSHELL_HOSTS) if path]


def controlled_environment(**overrides: str) -> dict[str, str]:
    environment = {key: value for key, value in os.environ.items() if key.upper() not in CONTROLLED_VARIABLES}
    environment.update(overrides)
    return environment


class PowerShell7AutoInstallTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-pwsh7-")
        cls.root = Path(cls._directory.name)
        cls.harness = cls.root / "harness.ps1"
        cls.harness.write_text(HARNESS, encoding="utf-8-sig")
        cls.fake_pwsh = cls.root / "pwsh.exe"
        cls.fake_pwsh.write_bytes(b"")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def workspace(self, host: str, name: str) -> Path:
        path = self.root / f"{Path(host).stem}-{name}"
        path.mkdir()
        (path / "winget-stub.cmd").write_text(WINGET_STUB, encoding="ascii")
        return path

    def command(self, host: str, state: Path, *, action: str = "render", winget: bool = True,
                pwsh: Path | None = None, edition: str = "Desktop", still_running: bool = False) -> list[str]:
        command = [host, "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                   "-File", str(self.harness), "-Bridge", str(BRIDGE), "-StateDirectory", str(state),
                   "-Action", action, "-Winget", str(state / "winget-stub.cmd") if winget else "",
                   "-Pwsh", str(pwsh or ""), "-Edition", edition]
        if still_running:
            command.append("-WingetStillRunning")
        return command

    def check(self, host: str, state: Path, environment: dict[str, str] | None = None, **options) -> dict | None:
        completed = subprocess.run(
            self.command(host, state, **options), capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=HOST_TIMEOUT_SECONDS, env=environment or controlled_environment(), check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        return json.loads(completed.stdout.strip().splitlines()[-1])["notice"]

    def state(self, state: Path) -> dict:
        return json.loads((state / "pwsh-install.json").read_text(encoding="utf-8-sig"))

    def wait_for_stub(self, state: Path) -> str:
        done = state / "winget-done.txt"
        deadline = time.monotonic() + STUB_WAIT_SECONDS
        while not done.exists() and time.monotonic() < deadline:
            time.sleep(STUB_POLL_SECONDS)
        self.assertTrue(done.exists(), "the winget stub never finished")
        return (state / "winget-arguments.txt").read_text(encoding="ascii", errors="replace").strip()

    def test_failed_install_is_reported_once(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "failed")
                notice = self.check(host, state)
                self.assertEqual(notice["status"], "installing")
                self.assertTrue(notice["relay_to_user"])
                self.assertIn("agreements were accepted on the user's behalf", notice["message"])
                arguments = self.wait_for_stub(state)
                self.assertTrue(arguments.startswith(EXPECTED_WINGET_ARGUMENTS), arguments)
                self.assertTrue(arguments.endswith(f'"{state / "pwsh-install.log"}"'), arguments)
                self.assertEqual(self.state(state)["state"], "installing")
                notice = self.check(host, state)
                self.assertEqual(notice["status"], "failed")
                self.assertIn("stub winget finished", notice["message"])
                self.assertEqual(self.state(state)["state"], "failed")
                self.assertIsNone(self.check(host, state))

    def test_successful_install_is_reported_once(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "succeeded")
                self.assertEqual(self.check(host, state)["status"], "installing")
                self.wait_for_stub(state)
                notice = self.check(host, state, pwsh=self.fake_pwsh, edition="Core")
                self.assertEqual(notice["status"], "succeeded")
                self.assertIn(str(self.fake_pwsh), notice["message"])
                self.assertEqual(self.state(state)["state"], "succeeded")
                self.assertIsNone(self.check(host, state, pwsh=self.fake_pwsh, edition="Core"))

    def test_running_install_stays_silent(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "running")
                self.assertEqual(self.check(host, state)["status"], "installing")
                self.wait_for_stub(state)
                self.assertIsNone(self.check(host, state, still_running=True))
                self.assertEqual(self.state(state)["state"], "installing")

    def test_install_does_not_hold_the_caller(self) -> None:
        # A child that inherits the caller's pipes keeps them open until it exits; the caller
        # must get its answer long before the slow stub finishes.
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "detached")
                (state / "slow").write_bytes(b"")
                started = time.monotonic()
                completed = subprocess.run(
                    self.command(host, state), capture_output=True, text=True, encoding="utf-8",
                    errors="replace", timeout=HOST_TIMEOUT_SECONDS, env=controlled_environment(), check=False,
                )
                elapsed = time.monotonic() - started
                self.assertEqual(completed.returncode, 0, completed.stderr)
                self.assertLess(elapsed, CALLER_BUDGET_SECONDS, f"the caller waited {elapsed:.1f}s for the stub")
                self.assertFalse((state / "winget-done.txt").exists())
                self.wait_for_stub(state)

    def test_read_only_actions_announce_without_installing(self) -> None:
        for host in self.hosts:
            for action in ("probe", "list-styles", "validate"):
                with self.subTest(host=Path(host).name, action=action):
                    state = self.workspace(host, f"announce-{action}")
                    notice = self.check(host, state, action=action)
                    self.assertEqual(notice["status"], "pending")
                    self.assertIn("ENDNOTE_WORD_AUTO_INSTALL_PWSH=0", notice["message"])
                    self.assertFalse((state / "pwsh-install.json").exists())
                    self.assertFalse((state / "winget-arguments.txt").exists())

    def test_missing_winget_is_reported_once(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "unavailable")
                self.assertEqual(self.check(host, state, action="probe", winget=False)["status"], "unavailable")
                self.assertFalse((state / "pwsh-install.json").exists())
                self.assertEqual(self.check(host, state, winget=False)["status"], "unavailable")
                self.assertEqual(self.state(state)["state"], "unavailable")
                self.assertIsNone(self.check(host, state, winget=False))
                # Once winget shows up, the recorded 'unavailable' does not block the install.
                self.assertEqual(self.check(host, state)["status"], "installing")
                self.wait_for_stub(state)

    def test_unreadable_state_is_set_aside(self) -> None:
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                state = self.workspace(host, "corrupt")
                (state / "pwsh-install.json").write_bytes(b"{ not json")
                self.assertEqual(self.check(host, state)["status"], "installing")
                self.assertTrue((state / "pwsh-install.json.bad").exists())
                self.wait_for_stub(state)

    def test_no_install_is_started_when_not_needed_or_not_wanted(self) -> None:
        cases = {
            "opt-out": ({}, {"ENDNOTE_WORD_AUTO_INSTALL_PWSH": "0"}),
            "pinned-host": ({}, {"ENDNOTE_WORD_POWERSHELL": "powershell.exe"}),
            "already-on-pwsh7": ({"edition": "Core"}, {}),
            "pwsh-present": ({"pwsh": self.fake_pwsh}, {}),
        }
        for host in self.hosts:
            for name, (options, variables) in cases.items():
                with self.subTest(host=Path(host).name, case=name):
                    state = self.workspace(host, name)
                    self.assertIsNone(self.check(host, state, controlled_environment(**variables), **options))
                    self.assertFalse((state / "pwsh-install.json").exists())
                    self.assertFalse((state / "winget-arguments.txt").exists())

    def test_bridge_result_carries_the_notice_once(self) -> None:
        windows_powershell = shutil.which("powershell.exe")
        if not windows_powershell:
            self.skipTest("Windows PowerShell 5.1 is not available.")
        profile = self.root / "e2e-profile"
        for folder in ("AppData", "LocalAppData", "ProgramFiles"):
            (profile / folder).mkdir(parents=True)
        system32 = Path(os.environ.get("SystemRoot", "C:/Windows")) / "System32"
        # Hide every PowerShell 7 and winget location the bridge looks at. Windows variable names
        # are case-insensitive, so drop any existing spelling before adding the override, and a
        # 64-bit process rebuilds ProgramFiles from ProgramW6432 at start, so override both.
        overrides = {"APPDATA": profile / "AppData", "LOCALAPPDATA": profile / "LocalAppData",
                     "PROGRAMFILES": profile / "ProgramFiles", "PROGRAMW6432": profile / "ProgramFiles",
                     "PATH": system32}
        environment = {key: value for key, value in controlled_environment().items() if key.upper() not in overrides}
        environment.update({key: str(value) for key, value in overrides.items()})
        state_file = profile / "LocalAppData" / "EndNoteForWordAgent" / "pwsh-install.json"
        results = []
        for action in ("list-styles", "configure-defaults", "configure-defaults"):
            arguments = ["-Action", action] + (["-Style", "IEEE"] if action == "list-styles" else [])
            completed = subprocess.run(
                [windows_powershell, "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-File", str(BRIDGE), *arguments],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=HOST_TIMEOUT_SECONDS, env=environment, check=False,
            )
            lines = [line for line in completed.stdout.splitlines() if line.strip()]
            self.assertTrue(lines, completed.stderr)
            results.append((json.loads(lines[-1]).get("powershell_notice"), completed.stderr, state_file.exists()))
        (read_only, read_only_stderr, state_after_read_only), (first, _, _), (second, second_stderr, _) = results
        self.assertEqual(read_only["status"], "unavailable")
        self.assertIn("POWERSHELL_NOTICE=", read_only_stderr)
        self.assertFalse(state_after_read_only)
        self.assertEqual(first["status"], "unavailable")
        self.assertEqual(json.loads(state_file.read_text(encoding="utf-8"))["state"], "unavailable")
        self.assertIsNone(second)
        self.assertNotIn("POWERSHELL_NOTICE=", second_stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
