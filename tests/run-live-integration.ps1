# Live end-to-end check of scripts\endnote-word.ps1 against the installed Word and EndNote.
#
# Requires ENDNOTE_TEST_LIBRARY to point at a real .enl library whose record numbers match
# tests\fixtures\live-manifest.json. The library is only ever read. The draft DOCX is built
# in a temporary directory and removed again unless -KeepArtifacts is passed.
#
# This file must parse under Windows PowerShell 5.1 as well as PowerShell 7, so it uses no
# ternary, null-coalescing or null-conditional operators.

[CmdletBinding()]
param(
    [switch]$KeepArtifacts,
    [string]$TargetStyle
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Force UTF-8 so the JSON report survives a cp950 console. 5.1 can throw when stdout is
# redirected, so the failure is reported and tolerated rather than allowed to abort the run.
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
    [Console]::Error.WriteLine("[live] Could not switch the console to UTF-8: $($_.Exception.Message)")
}

$Root = Split-Path -Parent $PSScriptRoot
$Bridge = Join-Path $Root 'scripts\endnote-word.ps1'
$FixtureManifest = Join-Path $PSScriptRoot 'fixtures\live-manifest.json'
$LibraryEnvironmentVariable = 'ENDNOTE_TEST_LIBRARY'
$TargetStyleEnvironmentVariable = 'ENDNOTE_TEST_TARGET_STYLE'
$PowerShellHosts = @('pwsh.exe', 'powershell.exe')
$DefaultTargetStyle = 'IEEE'
$JsonDepth = 20

# Word COM constants.
$WdFormatDocumentDefault = 16
$WdDoNotSaveChanges = 0

function Write-Status {
    param([Parameter(Mandatory = $true)][string]$Message)
    [Console]::Error.WriteLine("[live] $Message")
}

function Resolve-PowerShellHost {
    <#
        .SYNOPSIS
            Return the first installed PowerShell executable, preferring PowerShell 7.
    #>
    foreach ($candidate in $PowerShellHosts) {
        $command = Get-Command -Name $candidate -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($null -ne $command) { return $command.Source }
    }
    throw 'Neither pwsh.exe nor powershell.exe was found on PATH.'
}

function Invoke-Bridge {
    <#
        .SYNOPSIS
            Run one bridge action and return its single JSON result line as an object.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$HostExecutable,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    $output = & $HostExecutable -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass `
        -File $Bridge @Arguments
    $lines = @($output | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) {
        throw ('The bridge produced no output for: ' + ($Arguments -join ' '))
    }
    try {
        return ($lines[-1] | ConvertFrom-Json)
    }
    catch {
        throw ('The bridge did not return JSON for: ' + ($Arguments -join ' ') +
            ' Last line was: ' + $lines[-1])
    }
}

function New-DraftDocument {
    <#
        .SYNOPSIS
            Create the draft DOCX with one marker per manifest citation, through Word itself.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$Markers
    )
    $paragraphs = New-Object System.Collections.Generic.List[string]
    $paragraphs.Add('Live integration draft for the endnote-for-word bridge.')
    foreach ($marker in $Markers) {
        $paragraphs.Add('This sentence cites ' + $marker + ' and then ends.')
    }
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Add()
        $document.Content.Text = ($paragraphs -join "`r")
        $document.SaveAs2($Path, $WdFormatDocumentDefault)
    }
    finally {
        if ($null -ne $document) {
            try { $document.Close($WdDoNotSaveChanges) } catch { Write-Status "Could not close the draft: $($_.Exception.Message)" }
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
        if ($null -ne $word) {
            try { $word.Quit($WdDoNotSaveChanges) } catch { Write-Status "Could not quit Word: $($_.Exception.Message)" }
            $null = [Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$library = [Environment]::GetEnvironmentVariable($LibraryEnvironmentVariable)
if ([string]::IsNullOrWhiteSpace($library)) {
    throw "$LibraryEnvironmentVariable is not set, so the live integration run cannot start."
}
if (-not (Test-Path -LiteralPath $library -PathType Leaf)) {
    throw "$LibraryEnvironmentVariable does not point at a file: $library"
}
$libraryData = [System.IO.Path]::ChangeExtension($library, '.Data')
if (-not (Test-Path -LiteralPath $libraryData -PathType Container)) {
    throw "The library data directory is missing: $libraryData"
}
if (-not (Test-Path -LiteralPath $Bridge -PathType Leaf)) {
    throw "The bridge script was not found at $Bridge."
}

if ([string]::IsNullOrWhiteSpace($TargetStyle)) {
    $TargetStyle = [Environment]::GetEnvironmentVariable($TargetStyleEnvironmentVariable)
}
if ([string]::IsNullOrWhiteSpace($TargetStyle)) {
    $TargetStyle = $DefaultTargetStyle
}

$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('endnote-live-' + [Guid]::NewGuid().ToString('N'))
$baselineWordPids = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    New-Item -ItemType Directory -Path $testDirectory | Out-Null
    $inputPath = Join-Path $testDirectory 'input.docx'
    $outputPath = Join-Path $testDirectory 'output.docx'
    $restyledPath = Join-Path $testDirectory 'restyled.docx'
    $manifestPath = Join-Path $testDirectory 'live-manifest.json'

    $manifest = Get-Content -LiteralPath $FixtureManifest -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifest.library.path = $library
    $manifest | ConvertTo-Json -Depth $JsonDepth | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $markers = @($manifest.citations | ForEach-Object { $_.marker })
    $expectedCitations = $markers.Count
    Write-Status "Manifest has $expectedCitations citation(s); style '$($manifest.style)'."

    $powerShellHost = Resolve-PowerShellHost
    Write-Status "Bridge host: $powerShellHost"

    $probe = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @('-Action', 'probe', '-LibraryPath', $library)
    if (-not $probe.ok) { throw "Probe reported not ready: $($probe | ConvertTo-Json -Depth $JsonDepth -Compress)" }
    foreach ($check in @('word_com', 'endnote_executable', 'cwyw_template_loaded', 'cwyw_macro_callable', 'library_exists', 'library_data_exists')) {
        if (-not $probe.checks.$check) { throw "Probe check '$check' failed." }
    }

    $styleList = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @('-Action', 'list-styles', '-Style', $manifest.style)
    if (-not $styleList.ok) { throw 'list-styles failed for the manifest style.' }
    $installed = @($styleList.styles | ForEach-Object { $_.ToLowerInvariant() })
    if ($installed -notcontains $manifest.style.ToLowerInvariant()) {
        throw "The manifest style '$($manifest.style)' is not installed; rendering it would open a modal dialog."
    }

    New-DraftDocument -Path $inputPath -Markers $markers
    if (-not (Test-Path -LiteralPath $inputPath)) { throw 'The draft DOCX was not created.' }

    $render = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @(
        '-Action', 'render', '-InputPath', $inputPath, '-OutputPath', $outputPath, '-ManifestPath', $manifestPath)
    if (-not $render.ok) { throw "Render failed: $($render.error)" }
    if (-not (Test-Path -LiteralPath $outputPath)) { throw 'Render reported success but wrote no file.' }

    $validation = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @(
        '-Action', 'validate', '-InputPath', $outputPath, '-ManifestPath', $manifestPath)
    if (-not $validation.ok) { throw "Validation failed: $($validation.errors -join '; ')" }
    if ($validation.counts.native_citation_fields -ne $expectedCitations) {
        throw "Expected $expectedCitations citation field(s), found $($validation.counts.native_citation_fields)."
    }
    if ($validation.counts.bibliography_fields -ne 1) {
        throw "Expected exactly one bibliography field, found $($validation.counts.bibliography_fields)."
    }

    $restyle = $null
    $revalidation = $null
    $targetStyleList = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @('-Action', 'list-styles', '-Style', $TargetStyle)
    if (-not $targetStyleList.ok) { throw "list-styles failed for the target style '$TargetStyle'." }
    $targetInstalled = @($targetStyleList.styles | ForEach-Object { $_.ToLowerInvariant() })
    if ($targetInstalled -contains $TargetStyle.ToLowerInvariant()) {
        $restyle = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @(
            '-Action', 'set-style', '-InputPath', $outputPath, '-OutputPath', $restyledPath, '-Style', $TargetStyle)
        if (-not $restyle.ok) { throw "set-style failed: $($restyle.error)" }
        if ($restyle.style -ne $TargetStyle) { throw "set-style reported style '$($restyle.style)'." }

        $manifest.style = $TargetStyle
        $manifest | ConvertTo-Json -Depth $JsonDepth | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        $revalidation = Invoke-Bridge -HostExecutable $powerShellHost -Arguments @(
            '-Action', 'validate', '-InputPath', $restyledPath, '-ManifestPath', $manifestPath)
        if (-not $revalidation.ok) { throw "Validation after the style switch failed: $($revalidation.errors -join '; ')" }
        if ($revalidation.counts.native_citation_fields -ne $expectedCitations) {
            throw 'The style switch changed the citation field count.'
        }
    }
    else {
        Write-Status "Target style '$TargetStyle' is not installed; the style switch step was skipped."
    }

    [ordered]@{
        ok            = $true
        library       = $library
        input         = $inputPath
        output        = $outputPath
        restyled      = $restyledPath
        citations     = $expectedCitations
        probe         = $probe
        render        = $render
        validation    = $validation
        restyle       = $restyle
        revalidation  = $revalidation
    } | ConvertTo-Json -Depth $JsonDepth
}
catch {
    Write-Status "FAILED: $($_.Exception.Message)"
    throw
}
finally {
    $newWordProcesses = @(Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue |
        Where-Object { $baselineWordPids -notcontains $_.Id })
    if ($newWordProcesses.Count -gt 0) {
        Write-Status ('Test-created Word process(es) remain; not terminating automatically: ' +
            (($newWordProcesses | ForEach-Object { $_.Id }) -join ', '))
    }
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $testDirectory)) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force -ErrorAction SilentlyContinue
    }
    elseif ($KeepArtifacts) {
        Write-Status "Kept test artifacts: $testDirectory"
    }
}
