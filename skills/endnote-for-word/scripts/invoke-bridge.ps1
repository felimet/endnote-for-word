# Thin forwarder from the skill to the endnote-for-word COM bridge.
#
# The skill ships without the bridge, so this locates scripts/endnote-word.ps1 in the
# installed plugin (explicit ENDNOTE_WORD_PLUGIN_ROOT first, then the layout relative to
# this file) and runs it in a PowerShell host that can drive Word over COM. PowerShell 7 is
# preferred, Windows PowerShell 5.1 is accepted; the bridge parses and runs under both.
#
# All remaining arguments are passed straight through, for example:
#   ./invoke-bridge.ps1 -Action probe
#   ./invoke-bridge.ps1 -Action render -InputPath in.docx -OutputPath out.docx -ManifestPath m.json
#
# This file must parse under Windows PowerShell 5.1 as well as PowerShell 7, so it uses no
# ternary, null-coalescing or null-conditional operators.

[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(ValueFromRemainingArguments)]
    [string[]]$BridgeArguments
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# PowerShell decodes a child process's stdout with [Console]::OutputEncoding before relaying it.
# That defaults to the legacy code page here (cp950), which silently corrupts accented surnames
# and CJK titles in the bridge's single JSON result line. Forcing UTF-8 makes the relay a
# byte-for-byte passthrough. 5.1 can throw when stdout is redirected, so the assignment is guarded.
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
    [Console]::Error.WriteLine("[endnote-for-word] WARNING: could not force UTF-8 output: $($_.Exception.Message)")
}

# Ordered preference: PowerShell 7 first, Windows PowerShell 5.1 as the accepted fallback.
$PowerShellCandidates = @('pwsh.exe', 'powershell.exe')
$PowerShellEnvironmentVariable = 'ENDNOTE_WORD_POWERSHELL'
$PluginRootEnvironmentVariable = 'ENDNOTE_WORD_PLUGIN_ROOT'
$PowerShellInstallUrl = 'https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows'
$BridgeRelativePath = 'scripts\endnote-word.ps1'
# From skills/endnote-for-word/scripts/ back up to the repository root.
$RepositoryRootRelativePath = '..\..\..'

function Write-BridgeError {
    param([Parameter(Mandatory = $true)][string]$Message)
    [Console]::Error.WriteLine("[endnote-for-word] ERROR: $Message")
}

function Resolve-PowerShellHost {
    <#
        .SYNOPSIS
            Return the PowerShell executable used to run the bridge.
        .DESCRIPTION
            An explicit ENDNOTE_WORD_POWERSHELL override wins, then pwsh.exe, then
            powershell.exe. Throws an actionable message when none of them exist, because
            silently downgrading to some other shell would leave Word automation broken in a
            way that is hard to diagnose.
    #>
    $override = $env:ENDNOTE_WORD_POWERSHELL
    if (-not [string]::IsNullOrWhiteSpace($override)) {
        return $override
    }
    foreach ($candidate in $PowerShellCandidates) {
        $command = $null
        try {
            $command = Get-Command -Name $candidate -CommandType Application -ErrorAction Stop |
                Select-Object -First 1
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            # Expected when the candidate is not installed; the fixed locations are tried next.
            $command = $null
        }
        catch {
            Write-BridgeError "Could not probe '$candidate': $($_.Exception.Message)"
            $command = $null
        }
        if ($null -ne $command) {
            return $command.Source
        }
        if ($candidate -eq 'pwsh.exe') {
            # A PowerShell 7 that the bridge installed may not be on this shell's PATH yet: the
            # per-user MSIX app execution alias first, then the MSI location.
            $knownPaths = @()
            if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { $knownPaths += Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\pwsh.exe' }
            if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) { $knownPaths += Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe' }
            foreach ($knownPath in $knownPaths) {
                if (Test-Path -LiteralPath $knownPath -PathType Leaf) { return $knownPath }
            }
        }
    }
    throw ("Neither pwsh.exe (PowerShell 7) nor powershell.exe (Windows PowerShell 5.1) was " +
        "found on PATH. EndNote automation drives Word over COM and needs one of them. " +
        "Install PowerShell 7 from $PowerShellInstallUrl, or set " +
        "$PowerShellEnvironmentVariable to the full path of a PowerShell executable, then " +
        "retry. Do not substitute another shell.")
}

function Resolve-BridgeScript {
    <#
        .SYNOPSIS
            Return the path of scripts/endnote-word.ps1 from the installed plugin.
    #>
    $candidates = New-Object System.Collections.Generic.List[string]
    $pluginRoot = $env:ENDNOTE_WORD_PLUGIN_ROOT
    if (-not [string]::IsNullOrWhiteSpace($pluginRoot)) {
        $candidates.Add((Join-Path $pluginRoot $BridgeRelativePath))
    }
    $repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $RepositoryRootRelativePath))
    $candidates.Add((Join-Path $repositoryRoot $BridgeRelativePath))
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }
    throw ("The full endnote-for-word plugin is not installed: no bridge script was found at " +
        ($candidates -join '; ') + ". Register its MCP server, or set " +
        "$PluginRootEnvironmentVariable to the plugin root directory.")
}

try {
    $bridge = Resolve-BridgeScript
    $powerShellHost = Resolve-PowerShellHost
    if ($null -eq $BridgeArguments) {
        $BridgeArguments = @()
    }
    & $powerShellHost -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridge @BridgeArguments
    exit $LASTEXITCODE
}
catch {
    Write-BridgeError $_.Exception.Message
    exit 1
}
