# Launcher for the endnote-for-word stdio MCP server.
#
# Responsibilities:
#   1. Find a Python 3 interpreter to host scripts/mcp_server.py.
#   2. Check that a PowerShell host for the COM bridge exists, so a missing one fails at launch.
#      PowerShell 7 is preferred, Windows PowerShell 5.1 is accepted. The server resolves the
#      host again on every call, so a PowerShell 7 that the bridge installs in the background
#      is used without a restart; only an explicit ENDNOTE_WORD_POWERSHELL pins the host.
#
# This file must parse under Windows PowerShell 5.1 as well as PowerShell 7, so it uses no
# ternary, null-coalescing or null-conditional operators.

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# PowerShell decodes a child process's stdout with [Console]::OutputEncoding before relaying it.
# That defaults to the legacy code page here (cp950), which corrupts every non-ASCII character
# in the server's JSON-RPC frames. Forcing UTF-8 makes the relay a byte-for-byte passthrough.
# 5.1 can throw when stdout is redirected, so the assignment is guarded.
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
    [Console]::Error.WriteLine("[endnote-for-word] WARNING: could not force UTF-8 output: $($_.Exception.Message)")
}

# Ordered preference: PowerShell 7 first, Windows PowerShell 5.1 as the accepted fallback.
$PowerShellCandidates = @('pwsh.exe', 'powershell.exe')
$PythonCandidates = @('python.exe', 'python3.exe', 'python', 'python3')
$PowerShellEnvironmentVariable = 'ENDNOTE_WORD_POWERSHELL'
$PythonEnvironmentVariable = 'ENDNOTE_WORD_PYTHON'
$PowerShellInstallUrl = 'https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows'
$PythonInstallUrl = 'https://www.python.org/downloads/windows/'
$ServerFileName = 'mcp_server.py'

function Write-LauncherError {
    param([Parameter(Mandatory = $true)][string]$Message)
    [Console]::Error.WriteLine("[endnote-for-word] ERROR: $Message")
}

function Resolve-Executable {
    <#
        .SYNOPSIS
            Return the full path of the first candidate found on PATH, or $null when none exist.
    #>
    param(
        [Parameter(Mandatory = $true)][string[]]$Candidates
    )
    foreach ($candidate in $Candidates) {
        $command = $null
        try {
            $command = Get-Command -Name $candidate -CommandType Application -ErrorAction Stop |
                Select-Object -First 1
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            # Expected when the candidate is simply not installed; try the next one.
            continue
        }
        catch {
            Write-LauncherError "Could not probe '$candidate': $($_.Exception.Message)"
            continue
        }
        if ($null -ne $command) {
            return $command.Source
        }
    }
    return $null
}

try {
    $server = Join-Path $PSScriptRoot $ServerFileName
    if (-not (Test-Path -LiteralPath $server -PathType Leaf)) {
        throw "The MCP server script was not found at '$server'. The plugin installation is incomplete."
    }

    # The bridge host is checked here so a missing PowerShell fails loudly at launch instead
    # of surfacing later as an opaque tool error. It is deliberately not published: pinning a
    # Windows PowerShell 5.1 path here would keep the server on 5.1 after PowerShell 7 arrives.
    if ([string]::IsNullOrWhiteSpace($env:ENDNOTE_WORD_POWERSHELL)) {
        $bridgeHost = Resolve-Executable -Candidates $PowerShellCandidates
        if ([string]::IsNullOrWhiteSpace($bridgeHost)) {
            throw ("Neither pwsh.exe (PowerShell 7) nor powershell.exe (Windows PowerShell 5.1) " +
                "was found on PATH. The EndNote COM bridge cannot run without one. Install " +
                "PowerShell 7 from $PowerShellInstallUrl, or set $PowerShellEnvironmentVariable " +
                "to the full path of a PowerShell executable, then start the server again.")
        }
    }

    $python = $env:ENDNOTE_WORD_PYTHON
    if ([string]::IsNullOrWhiteSpace($python)) {
        $python = Resolve-Executable -Candidates $PythonCandidates
    }
    if ([string]::IsNullOrWhiteSpace($python)) {
        throw ("Python 3 was not found on PATH. Install it from $PythonInstallUrl, or set " +
            "$PythonEnvironmentVariable to the full path of a python.exe, then start the " +
            "server again.")
    }
    Set-Item -Path "env:$PythonEnvironmentVariable" -Value $python

    & $python $server
    exit $LASTEXITCODE
}
catch {
    Write-LauncherError $_.Exception.Message
    exit 1
}
