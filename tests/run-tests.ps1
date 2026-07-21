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
$CleanupLibrary = Join-Path $Root 'scripts\endnote_cleanup.ps1'
if (-not (Test-Path -LiteralPath $CleanupLibrary -PathType Leaf)) {
    throw "Cleanup library is missing: $CleanupLibrary"
}
. $CleanupLibrary

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

function Test-CleanupHelpers {
    param([Parameter(Mandatory = $true)][string]$TestRoot)

    $checkDirectory = Join-Path $TestRoot 'cleanup-helper-check'
    New-Item -ItemType Directory -Path $checkDirectory -ErrorAction Stop | Out-Null
    $destination = Join-Path $checkDirectory 'output.docx'
    $staleFile = Join-Path $checkDirectory ('.output.' + [Guid]::NewGuid().ToString('N') + '.tmp.docx')
    $recentFile = Join-Path $checkDirectory ('.output.' + [Guid]::NewGuid().ToString('N') + '.work.docx')
    $unrelatedFile = Join-Path $checkDirectory 'keep-me.tmp.docx'

    Set-Content -LiteralPath $staleFile -Value 'stale' -Encoding ASCII
    Set-Content -LiteralPath $recentFile -Value 'recent' -Encoding ASCII
    Set-Content -LiteralPath $unrelatedFile -Value 'unrelated' -Encoding ASCII
    (Get-Item -LiteralPath $staleFile -Force).LastWriteTimeUtc = [DateTime]::UtcNow.AddHours(-25)

    $removed = @(Remove-EndNoteStaleDocumentFiles -Destination $destination -OlderThanHours 24)
    if ($removed.Count -ne 1 -or -not (Test-EndNoteSamePath -Left $removed[0] -Right $staleFile)) {
        throw 'The stale document sweep did not report exactly the owned stale artifact.'
    }
    if (Test-Path -LiteralPath $staleFile) {
        throw 'The stale document sweep reported success but left the stale artifact behind.'
    }
    if (-not (Test-Path -LiteralPath $recentFile) -or -not (Test-Path -LiteralPath $unrelatedFile)) {
        throw 'The stale document sweep removed a recent or unrelated file.'
    }
    if (-not (Remove-EndNoteGeneratedDocumentFile -Path $recentFile -Destination $destination)) {
        throw 'The exact generated-file cleanup did not remove its owned artifact.'
    }
    if (Test-Path -LiteralPath $recentFile) {
        throw 'The exact generated-file cleanup did not verify removal.'
    }
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
$sourceRoots = @('scripts', 'tests', 'skills')
$powerShellFiles = New-Object System.Collections.Generic.List[System.IO.FileInfo]
foreach ($sourceRoot in $sourceRoots) {
    $sourcePath = Join-Path $Root $sourceRoot
    if (Test-Path -LiteralPath $sourcePath -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $sourcePath -Recurse -Filter '*.ps1' -File) {
            $powerShellFiles.Add($file)
        }
    }
}
foreach ($file in $powerShellFiles) {
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

$tempParent = [System.IO.Path]::GetTempPath()
$testTempRoot = Join-Path $tempParent ('endnote-tests-' + [Guid]::NewGuid().ToString('N'))
$parseCheckPath = Join-Path $testTempRoot 'parse-check.ps1'
$previousLiveSetting = [Environment]::GetEnvironmentVariable($LiveTestEnvironmentVariable)
$previousBytecodeSetting = [Environment]::GetEnvironmentVariable($BytecodeEnvironmentVariable)
$previousTempSetting = [Environment]::GetEnvironmentVariable('TEMP')
$previousTmpSetting = [Environment]::GetEnvironmentVariable('TMP')
$checkedHosts = New-Object System.Collections.Generic.List[string]
$staleTestDirectoriesRemoved = New-Object System.Collections.Generic.List[string]
$cleanupCompleted = $false

function Restore-TestEnvironment {
    [Environment]::SetEnvironmentVariable($LiveTestEnvironmentVariable, $previousLiveSetting, 'Process')
    [Environment]::SetEnvironmentVariable($BytecodeEnvironmentVariable, $previousBytecodeSetting, 'Process')
    [Environment]::SetEnvironmentVariable('TEMP', $previousTempSetting, 'Process')
    [Environment]::SetEnvironmentVariable('TMP', $previousTmpSetting, 'Process')
}

function Complete-TestCleanup {
    Remove-EndNoteRepoCacheDirectory -RepositoryRoot $Root -RelativePath 'scripts\__pycache__' | Out-Null
    Remove-EndNoteRepoCacheDirectory -RepositoryRoot $Root -RelativePath 'tests\__pycache__' | Out-Null
    Remove-EndNoteGeneratedDirectory -Path $testTempRoot -Parent $tempParent -Prefix 'endnote-tests' | Out-Null
}

try {
    foreach ($staleDirectory in @(Remove-EndNoteStaleTestDirectories -Parent $tempParent -Prefix 'endnote-tests')) {
        $staleTestDirectoriesRemoved.Add($staleDirectory) | Out-Null
        Write-Status "Removed stale test directory: $staleDirectory"
    }
    New-Item -ItemType Directory -Path $testTempRoot -ErrorAction Stop | Out-Null
    Set-Item -Path 'env:TEMP' -Value $testTempRoot
    Set-Item -Path 'env:TMP' -Value $testTempRoot
    Test-CleanupHelpers -TestRoot $testTempRoot
    Write-Status 'Cleanup ownership, stale-age and removal-verification checks passed.'

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

    Restore-TestEnvironment
    Complete-TestCleanup
    $cleanupCompleted = $true

    [ordered]@{
        ok                = $true
        python            = $python
        python_tests      = 'passed'
        powershell_parse  = 'passed'
        parse_check_hosts = @($checkedHosts)
        live_integration  = $liveIntegration
        cleanup           = 'verified'
        stale_removed     = @($staleTestDirectoriesRemoved)
    } | ConvertTo-Json -Depth 5 -Compress
}
catch {
    Write-Status "FAILED: $($_.Exception.Message)"
    throw
}
finally {
    if (-not $cleanupCompleted) {
        Restore-TestEnvironment
        Complete-TestCleanup
    }
}
