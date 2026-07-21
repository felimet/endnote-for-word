# Offline test runner for endnote-for-word.
#
# Three layers, in order:
#   1. Python unittest discovery over tests\test_*.py.
#   2. A PowerShell parse check of every .ps1 in the repository, run under PowerShell 7 and
#      under Windows PowerShell 5.1 when both are installed. The 5.1 pass is what proves the
#      scripts really parse on the older host; the Python lint in test_static.py catches the
#      7-only operators that 5.1 would reject at runtime rather than at parse time.
#   3. The optional live integration run, only with -IncludeLive.
#
# This file must itself parse under Windows PowerShell 5.1, so it uses no ternary,
# null-coalescing or null-conditional operators.

[CmdletBinding()]
param(
    [switch]$IncludeLive,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The console on a zh-TW Windows install defaults to cp950, which mangles the JSON summary
# and the test output. Windows PowerShell 5.1 can throw here when stdout is redirected, so
# the failure is reported and tolerated rather than allowed to abort the run.
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
    [Console]::Error.WriteLine("[run-tests] Could not switch the console to UTF-8: $($_.Exception.Message)")
}

$Root = Split-Path -Parent $PSScriptRoot
$PythonEnvironmentVariable = 'ENDNOTE_WORD_PYTHON'
$LiveTestEnvironmentVariable = 'ENDNOTE_LIVE_TEST'
$BytecodeEnvironmentVariable = 'PYTHONDONTWRITEBYTECODE'
$PythonCandidates = @('python.exe', 'python3.exe', 'python', 'python3')
$PowerShellHosts = @('pwsh.exe', 'powershell.exe')
$TestFilePattern = 'test_*.py'

function Write-Status {
    param([Parameter(Mandatory = $true)][string]$Message)
    [Console]::Error.WriteLine("[run-tests] $Message")
}

function Resolve-Python {
    <#
        .SYNOPSIS
            Return a python executable that actually starts.
        .DESCRIPTION
            An explicit ENDNOTE_WORD_PYTHON always wins. Otherwise EVERY PATH match for every
            candidate name is launched with --version, not just the first one: a PATH entry can
            exist and still be blocked by an application control policy or be a Microsoft Store
            stub, and such a shim is often the first match, with the real interpreter later on
            PATH. Each executable is tried at most once because the candidate names overlap.
    #>
    $configured = [Environment]::GetEnvironmentVariable($PythonEnvironmentVariable)
    if (-not [string]::IsNullOrWhiteSpace($configured)) {
        return $configured
    }
    $tried = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in $PythonCandidates) {
        $commands = @(Get-Command -Name $candidate -CommandType Application -All -ErrorAction SilentlyContinue)
        foreach ($command in $commands) {
            if ($tried -contains $command.Source) { continue }
            $tried.Add($command.Source)
            try {
                $null = & $command.Source '--version' 2>&1
                if ($LASTEXITCODE -eq 0) { return $command.Source }
                Write-Status "Ignoring '$($command.Source)': it exited with $LASTEXITCODE."
            }
            catch {
                Write-Status "Ignoring '$($command.Source)': $($_.Exception.Message)"
            }
        }
    }
    throw ("No usable Python 3 interpreter was found. Install Python 3 or set " +
        "$PythonEnvironmentVariable to the full path of a working python.exe.")
}

$ParseCheckScript = @'
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Root)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$failures = New-Object System.Collections.Generic.List[string]
foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -Filter '*.ps1') {
    $tokens = $null
    $parseErrors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    if ($null -ne $parseErrors -and $parseErrors.Count -gt 0) {
        $failures.Add(($file.FullName + ': ' + (($parseErrors | ForEach-Object { $_.Message }) -join '; ')))
    }
}
if ($failures.Count -gt 0) {
    [Console]::Error.WriteLine('Parse failures:')
    foreach ($failure in $failures) { [Console]::Error.WriteLine('  ' + $failure) }
    exit 1
}
exit 0
'@

$parseCheckPath = Join-Path ([System.IO.Path]::GetTempPath()) ('endnote-parse-check-' + [Guid]::NewGuid().ToString('N') + '.ps1')
$previousLiveSetting = [Environment]::GetEnvironmentVariable($LiveTestEnvironmentVariable)
$previousBytecodeSetting = [Environment]::GetEnvironmentVariable($BytecodeEnvironmentVariable)
$checkedHosts = New-Object System.Collections.Generic.List[string]

try {
    $python = Resolve-Python
    Write-Status "Using python: $python"

    if ($IncludeLive) {
        Set-Item -Path "env:$LiveTestEnvironmentVariable" -Value '1'
    }
    Set-Item -Path "env:$BytecodeEnvironmentVariable" -Value '1'

    & $python -m unittest discover -s $PSScriptRoot -t $PSScriptRoot -p $TestFilePattern -v
    $pythonExitCode = $LASTEXITCODE
    if ($pythonExitCode -ne 0) {
        throw "Python tests failed with exit code $pythonExitCode."
    }

    Set-Content -LiteralPath $parseCheckPath -Value $ParseCheckScript -Encoding UTF8
    foreach ($hostName in $PowerShellHosts) {
        $hostCommand = Get-Command -Name $hostName -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($null -eq $hostCommand) {
            Write-Status "Skipping the parse check on '$hostName': not installed."
            continue
        }
        & $hostCommand.Source -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass `
            -File $parseCheckPath -Root $Root
        if ($LASTEXITCODE -ne 0) {
            throw "PowerShell parse check failed under $($hostCommand.Source)."
        }
        $checkedHosts.Add($hostCommand.Name)
        Write-Status "Parse check passed under $($hostCommand.Source)."
    }
    if ($checkedHosts.Count -eq 0) {
        throw 'No PowerShell host was available to parse-check the scripts.'
    }

    $liveIntegration = 'skipped'
    if ($IncludeLive) {
        & (Join-Path $PSScriptRoot 'run-live-integration.ps1') -KeepArtifacts:$KeepArtifacts
        if ($LASTEXITCODE -ne 0) {
            throw "Live integration run failed with exit code $LASTEXITCODE."
        }
        $liveIntegration = 'passed'
    }

    [ordered]@{
        ok                = $true
        python            = $python
        python_tests      = 'passed'
        powershell_parse  = 'passed'
        parse_check_hosts = @($checkedHosts)
        live_integration  = $liveIntegration
    } | ConvertTo-Json -Depth 5 -Compress
}
catch {
    Write-Status "FAILED: $($_.Exception.Message)"
    throw
}
finally {
    if ($null -eq $previousLiveSetting) {
        Remove-Item -Path "env:$LiveTestEnvironmentVariable" -ErrorAction SilentlyContinue
    }
    else {
        Set-Item -Path "env:$LiveTestEnvironmentVariable" -Value $previousLiveSetting
    }
    if ($null -eq $previousBytecodeSetting) {
        Remove-Item -Path "env:$BytecodeEnvironmentVariable" -ErrorAction SilentlyContinue
    }
    else {
        Set-Item -Path "env:$BytecodeEnvironmentVariable" -Value $previousBytecodeSetting
    }
    if (Test-Path -LiteralPath $parseCheckPath) {
        Remove-Item -LiteralPath $parseCheckPath -Force -ErrorAction SilentlyContinue
    }
}
