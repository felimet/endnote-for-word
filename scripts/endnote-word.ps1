[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet(
        'probe',
        'list-styles',
        'configure-defaults',
        'render',
        'validate',
        'update',
        'set-style',
        'unformat',
        'import-ris'
    )]
    [string]$Action,

    [string]$InputPath,
    [string]$OutputPath,
    [string]$ManifestPath,
    [string]$LibraryPath,
    [string]$RisPath,
    [string]$EmptyLibraryTemplatePath,
    [string]$Style,
    [switch]$Overwrite
)

# EndNote 21 Cite While You Write bridge for Microsoft Word.
#
# This file must parse and run under BOTH Windows PowerShell 5.1 and PowerShell 7.
# Therefore: no ternary operator, no null-coalescing, no null-conditional, no clean{} block,
# no Join-String, no ForEach-Object -Parallel. Use if/else, New-Object and [ordered]@{}.
#
# Contract:
#   - Exactly one compact JSON line on stdout, as the last stdout line.
#   - All logging goes to stderr through Write-Log.
#   - Exit codes: 0 ok, 1 unhandled error, 2 probe not ready, 3 postcondition failed.
#   - Source documents and libraries are user data. The input document is never opened for
#     writing. The import-ris action is the sole library-writing action and delegates every
#     write to EndNote's own Direct Export path; this script never edits sdb.eni directly.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$cleanupLibraryPath = Join-Path $PSScriptRoot 'endnote_cleanup.ps1'
if (-not (Test-Path -LiteralPath $cleanupLibraryPath -PathType Leaf)) {
    throw "Cleanup library is missing: $cleanupLibraryPath"
}
. $cleanupLibraryPath

# The console code page on a zh-TW Windows install defaults to cp950, which mangles the JSON
# result line whenever a style title or author name carries non-ASCII characters.
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
}
catch {
    # Windows PowerShell 5.1 can throw here when stdout is redirected. Surface it on stderr
    # instead of hiding it; the run itself can continue with the inherited encoding.
    [Console]::Error.WriteLine("[WARN] Could not switch console output encoding to UTF-8: $($_.Exception.Message)")
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,
        [Parameter(Mandatory)]
        [string]$Message
    )
    [Console]::Error.WriteLine("[$Level] $Message")
}

function Get-IntegerEnvironmentValue {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Default
    )
    $raw = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    $parsed = 0
    if (-not [int]::TryParse($raw.Trim(), [ref]$parsed) -or $parsed -le 0) {
        Write-Log -Level WARN -Message "Ignoring invalid $Name value '$raw'; falling back to $Default."
        return $Default
    }
    return $parsed
}

# ---------------------------------------------------------------------------------------------
# Constants. No magic numbers anywhere below this block.
# ---------------------------------------------------------------------------------------------

$script:PluginRoot = Split-Path -Parent $PSScriptRoot
$script:PackagedDefaultsPath = Join-Path $script:PluginRoot 'config\defaults.json'
$script:UserConfigDirectory = Join-Path $env:APPDATA 'EndNoteForWordAgent'
$script:UserDefaultsPath = Join-Path $script:UserConfigDirectory 'defaults.json'

# CWYW macro names. Every macro whose name ends in "Dlg" opens a modal dialog that blocks
# Application.Run forever, so Invoke-CwywMacro refuses to call one. ENRemoveFieldCodes is
# deliberately absent: it was measured to raise a modal confirmation dialog and never return
# under automation, so converting a document to plain text stays a manual EndNote ribbon
# operation and there is no remove-field-codes action.
$script:CwywFormatMacro = 'ENFormatBibliography'
$script:CwywUnformatMacro = 'ENUnformatCitations'
$script:CwywAuthorYearMacro = 'ENAuthorYearFormat'
$script:InteractiveMacroPattern = 'Dlg$'
$script:AllowedCwywMacros = @(
    $script:CwywFormatMacro,
    $script:CwywUnformatMacro,
    $script:CwywAuthorYearMacro
)

$script:CwywAddInProgId = 'EndNote.WordAddin.Connect'
$script:CwywTemplateFileName = 'EndNote Cwyw.dotm'
$script:CwywTemplateNamePattern = 'EndNote Cwyw*.dotm'
$script:CwywRelativePath = 'ResearchSoft\Cwyw'

$script:EndNoteExecutableName = 'EndNote.exe'
$script:EndNoteProcessName = 'EndNote'
$script:EndNoteDirectoryFilter = 'EndNote*'
$script:EndNoteStylesFolderName = 'Styles'
$script:EndNoteStyleExtension = '.ens'
$script:EndNoteStyleFilter = '*.ens'
$script:RisExtension = '.ris'
$script:PersonalStylesRelativePath = 'EndNote\Styles'
$script:LibraryDataDirectorySuffix = '.Data'
$script:EndNoteRegistryRoots = @(
    'HKLM:\SOFTWARE\WOW6432Node\ISI ResearchSoft\EndNote',
    'HKLM:\SOFTWARE\ISI ResearchSoft\EndNote',
    'HKCU:\Software\ISI ResearchSoft\EndNote'
)
$script:EndNoteRecentLibrariesRegistryPath = 'HKCU:\Software\ISI ResearchSoft\EndNote\Recent Libraries'
$script:EndNoteMostRecentLibraryValueName = 'File1'
$script:EndNoteAppPathsKeys = @(
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe'
)
$script:ProgIdRegistryPrefix = 'Registry::HKEY_CLASSES_ROOT\'

# EndNote document variables. Both are also persisted in word/settings.xml as w:docVar, so the
# offline structure tests can read exactly the same values without Word.
$script:LayoutVariableName = 'EN.Layout'
$script:LibrariesVariableName = 'EN.Libraries'
$script:LayoutStylePattern = '<Style>(.*?)</Style>'
$script:LayoutRootOpenTag = '<ENLayout>'
$script:LayoutStylePlaceholder = '@@ENDNOTE_STYLE@@'
# Verified EN.Layout payload written by EndNote 21 itself. It is only used when the document has
# no EN.Layout variable at all, which means EndNote has never formatted the document. EndNote
# rewrites the non-Style elements from the user's own preferences on the next format run; the
# delimiters must stay { and } because the temporary citation syntax depends on them.
$script:DefaultLayoutTemplate = '<ENLayout><Style>' + $script:LayoutStylePlaceholder + '</Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim><FontName>Times New Roman</FontName><FontSize>12</FontSize><ReflistTitle></ReflistTitle><StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent><HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter><HyperlinksEnabled>0</HyperlinksEnabled><HyperlinksVisible>0</HyperlinksVisible><EnableBibliographyCategories>0</EnableBibliographyCategories></ENLayout>'

# Field code classification. EN.CITE.DATA must be tested first because "ADDIN EN.CITE.DATA"
# also satisfies the EN.CITE pattern; that ordering is the citation counting rule:
# citation_count = count("ADDIN EN.CITE*") - count("ADDIN EN.CITE.DATA").
$script:CitationDataFieldPattern = '^\s*ADDIN\s+EN\.CITE\.DATA\b'
$script:CitationFieldPattern = '^\s*ADDIN\s+EN\.CITE\b'
$script:BibliographyFieldPattern = '^\s*ADDIN\s+EN\.REFLIST\b'
$script:HyperlinkFieldPattern = '^\s*HYPERLINK\b'

# Validation patterns.
$script:CitationMarkerPrefix = '{{CITE:'
$script:TemporaryCitationPattern = '\{[^{}]*#\d+[^{}]*\}'
$script:BrokenReferenceStrings = @('Error!', 'Reference source not found', '{Citation}', '{Bibliography}')
$script:RecordNumberElementFormat = '<rec-number>{0}</rec-number>'
$script:DoiPattern = '^10\.\d{4,9}/[!-~]+$'
$script:DoiResolverPrefix = 'https://doi.org/'
$script:DuplicateDoiResolverPattern = '(?i)https://doi\.org/(?:https?://(?:dx\.)?doi\.org/|doi:\s*)'

# Manifest schema v1.
$script:ManifestSchemaVersion = 1
$script:CitationIdPattern = '^[A-Za-z][A-Za-z0-9_-]*$'
$script:CitedPagesPattern = '^[0-9]+(\s*-\s*[0-9]+)?$'
# Characters that would break the generated temporary citation or trigger an interactive
# EndNote dialog (a leading backslash is the unsupported prefix / exclude-author syntax).
$script:ForbiddenCitationTextPattern = '[{}#@;\\]'
$script:AllowedManifestKeys = @('schema_version', 'library', 'style', 'citations', 'bibliography')
$script:AllowedCitationKeys = @('id', 'marker', 'targets', 'display', 'pages', 'suffix')
$script:AllowedTargetKeys = @('rec_number', 'author', 'year', 'doi')
$script:AllowedLibraryKeys = @('path')
$script:AllowedBibliographyKeys = @('expected')
$script:RejectedCitationKeys = @('prefix')
$script:DefaultCitationDisplay = 'default'
$script:AuthorYearCitationDisplay = 'author-year'
$script:AllowedCitationDisplays = @(
    $script:DefaultCitationDisplay,
    $script:AuthorYearCitationDisplay
)
$script:PythonEnvironmentVariable = 'ENDNOTE_WORD_PYTHON'
$script:PythonCandidates = @('python.exe', 'python3.exe', 'python', 'python3')
$script:LibraryScriptPath = Join-Path $PSScriptRoot 'endnote_library.py'

# Word enumeration values used by this bridge.
# WdSaveFormat.wdFormatXMLDocument. Not wdFormatDocumentDefault (16): that value honours the
# user's "Save files in this format" setting, so a machine configured for Word 97-2003 would
# write a binary .doc payload under the mandatory .docx extension.
$script:WdFormatXmlDocument = 12   # WdSaveFormat.wdFormatXMLDocument
$script:WdFindStop = 0             # WdFindWrap.wdFindStop
$script:WdDoNotSaveChanges = 0     # WdSaveOptions.wdDoNotSaveChanges
$script:WdSaveChanges = -1         # WdSaveOptions.wdSaveChanges
$script:WordAlertsNone = 0         # WdAlertLevel.wdAlertsNone
$script:WordProcessName = 'WINWORD'

$script:DocxExtension = '.docx'
$script:TemporaryDocumentSuffix = '.tmp.docx'
$script:WorkingCopySuffix = '.work.docx'
$script:MaxStyleResults = 200
$script:MaxStyleSuggestions = 10
$script:ZeroVersion = [version]'0.0'

$script:AutomationPollMilliseconds = 250
$script:AutomationTimeoutMilliseconds = Get-IntegerEnvironmentValue -Name 'ENDNOTE_AUTOMATION_TIMEOUT_MS' -Default 60000
$script:EndNoteReadyTimeoutMilliseconds = Get-IntegerEnvironmentValue -Name 'ENDNOTE_READY_TIMEOUT_MS' -Default 90000

# ---------------------------------------------------------------------------------------------
# Mutable script state.
# ---------------------------------------------------------------------------------------------

$script:Word = $null
$script:Document = $null
$script:Selection = $null
$script:WorkingCopyPath = $null
$script:WorkingCopyDestination = $null
$script:TemporarySavePath = $null
$script:TemporarySaveDestination = $null
$script:CleanupFailures = New-Object 'System.Collections.Generic.List[string]'
$script:CleanupRemoved = New-Object 'System.Collections.Generic.List[string]'
$script:StaleArtifactsRemoved = New-Object 'System.Collections.Generic.List[string]'
$script:EndNoteInstallationCache = $null
$script:CwywTemplateCache = $null
$script:InstalledStylesCache = $null

# ---------------------------------------------------------------------------------------------
# Generic helpers.
# ---------------------------------------------------------------------------------------------

function Write-JsonResult {
    param([Parameter(Mandatory)]$Value)
    $Value | ConvertTo-Json -Depth 20 -Compress
}

function Resolve-ExistingFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "$Label is required." }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label does not exist: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-OutputFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Extension = $script:DocxExtension
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'OutputPath is required.' }
    $fullPath = [IO.Path]::GetFullPath($Path)
    if ([IO.Path]::GetExtension($fullPath) -ine $Extension) {
        throw "OutputPath must use the $Extension extension: $fullPath"
    }
    $parent = Split-Path -Parent $fullPath
    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw "OutputPath must be an absolute path with a parent directory: $fullPath"
    }
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    if ((Test-Path -LiteralPath $fullPath) -and -not $Overwrite) {
        throw "OutputPath already exists. Pass -Overwrite to replace it: $fullPath"
    }
    return $fullPath
}

function Release-ComObject {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
        catch { Write-Log -Level WARN -Message "Could not release a COM object: $($_.Exception.Message)" }
    }
}

function Get-OptionalProperty {
    param(
        [Parameter(Mandatory)][AllowNull()]$Object,
        [Parameter(Mandatory)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-AgentDefaults {
    $path = $script:PackagedDefaultsPath
    if (Test-Path -LiteralPath $script:UserDefaultsPath -PathType Leaf) {
        $path = $script:UserDefaultsPath
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Agent defaults are missing: $path"
    }
    try { return Get-Content -Raw -LiteralPath $path | ConvertFrom-Json }
    catch { throw "Agent defaults are not valid JSON ($path): $($_.Exception.Message)" }
}

function Get-DefaultOutputStyle {
    $defaults = Get-AgentDefaults
    $styleSettings = Get-OptionalProperty -Object $defaults -Name 'style'
    $styleName = [string](Get-OptionalProperty -Object $styleSettings -Name 'default_output_style')
    if ([string]::IsNullOrWhiteSpace($styleName)) {
        throw 'Agent defaults do not define style.default_output_style.'
    }
    return $styleName.Trim()
}

function Resolve-PythonExecutable {
    $configured = [Environment]::GetEnvironmentVariable($script:PythonEnvironmentVariable)
    if (-not [string]::IsNullOrWhiteSpace($configured)) {
        if (-not (Test-Path -LiteralPath $configured -PathType Leaf)) {
            throw "$($script:PythonEnvironmentVariable) does not point to a Python executable: $configured"
        }
        return (Resolve-Path -LiteralPath $configured).Path
    }

    foreach ($candidate in $script:PythonCandidates) {
        try {
            $command = Get-Command -Name $candidate -CommandType Application -ErrorAction Stop |
                Select-Object -First 1
            if ($null -ne $command) { return $command.Source }
        }
        catch [System.Management.Automation.CommandNotFoundException] { continue }
        catch { Write-Log -Level WARN -Message "Could not probe Python candidate '$candidate': $($_.Exception.Message)" }
    }
    throw "Python 3 is required for DOI and library precheck. Set $($script:PythonEnvironmentVariable) to the full python.exe path."
}

function ConvertTo-ProcessArgument {
    param([Parameter(Mandatory)][string]$Value)
    if ($Value.Contains('"')) { throw "A child-process argument cannot contain a double quote: $Value" }
    return '"' + $Value + '"'
}

function Invoke-HiddenProcess {
    param(
        [Parameter(Mandatory)][string]$Executable,
        [Parameter(Mandatory)][string]$Arguments
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Executable
    $startInfo.Arguments = $Arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    try {
        $utf8 = New-Object System.Text.UTF8Encoding $false
        $startInfo.StandardOutputEncoding = $utf8
        $startInfo.StandardErrorEncoding = $utf8
    }
    catch { Write-Log -Level WARN -Message "Could not force UTF-8 for a child process: $($_.Exception.Message)" }

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) { throw "The child process did not start: $Executable" }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        return [ordered]@{ exit_code = $process.ExitCode; stdout = $stdout; stderr = $stderr }
    }
    finally { $process.Dispose() }
}

function Invoke-ManifestPrecheck {
    param(
        [Parameter(Mandatory)][string]$LibraryFile,
        [Parameter(Mandatory)][string]$ManifestFile
    )
    if (-not (Test-Path -LiteralPath $script:LibraryScriptPath -PathType Leaf)) {
        throw "The read-only library precheck script is missing: $script:LibraryScriptPath"
    }
    $python = Resolve-PythonExecutable
    $arguments = @(
        (ConvertTo-ProcessArgument -Value $script:LibraryScriptPath),
        'precheck',
        '--library',
        (ConvertTo-ProcessArgument -Value $LibraryFile),
        '--manifest',
        (ConvertTo-ProcessArgument -Value $ManifestFile)
    ) -join ' '

    $captured = Invoke-HiddenProcess -Executable $python -Arguments $arguments
    if (-not [string]::IsNullOrWhiteSpace($captured.stderr)) {
        Write-Log -Level INFO -Message $captured.stderr.Trim()
    }
    $lines = @($captured.stdout -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) { throw "The DOI and library precheck returned no JSON (exit $($captured.exit_code))." }
    try { $result = $lines[$lines.Count - 1] | ConvertFrom-Json }
    catch { throw "The DOI and library precheck returned invalid JSON: $($_.Exception.Message)" }
    if ($captured.exit_code -eq 0 -and [bool](Get-OptionalProperty -Object $result -Name 'ok')) {
        return $result
    }
    $details = New-Object System.Collections.Generic.List[string]
    foreach ($problem in @(Get-OptionalProperty -Object $result -Name 'problems')) {
        $details.Add(([string](Get-OptionalProperty -Object $problem -Name 'problem')) + ': ' + ([string](Get-OptionalProperty -Object $problem -Name 'message')))
    }
    $errorText = [string](Get-OptionalProperty -Object $result -Name 'error')
    if (-not [string]::IsNullOrWhiteSpace($errorText)) { $details.Add($errorText) }
    if ($details.Count -eq 0) { $details.Add("precheck exited $($captured.exit_code)") }
    throw ('DOI and library precheck failed: ' + ($details -join ' | '))
}

function Get-EndNoteRecordCount {
    param(
        [Parameter(Mandatory)][string]$LibraryFile,
        [switch]$IncludeTrash
    )
    if (-not (Test-Path -LiteralPath $script:LibraryScriptPath -PathType Leaf)) {
        throw "The read-only library counter is missing: $script:LibraryScriptPath"
    }
    $python = Resolve-PythonExecutable
    $arguments = @(
        (ConvertTo-ProcessArgument -Value $script:LibraryScriptPath),
        'count',
        '--library',
        (ConvertTo-ProcessArgument -Value $LibraryFile)
    )
    if ($IncludeTrash) { $arguments += '--include-trash' }
    $argumentText = $arguments -join ' '
    $captured = Invoke-HiddenProcess -Executable $python -Arguments $argumentText
    if (-not [string]::IsNullOrWhiteSpace($captured.stderr)) {
        Write-Log -Level INFO -Message $captured.stderr.Trim()
    }
    $lines = @($captured.stdout -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($lines.Count -eq 0) { throw "The library counter returned no JSON (exit $($captured.exit_code))." }
    try { $result = $lines[$lines.Count - 1] | ConvertFrom-Json }
    catch { throw "The library counter returned invalid JSON: $($_.Exception.Message)" }
    if ($captured.exit_code -ne 0 -or -not [bool](Get-OptionalProperty -Object $result -Name 'ok')) {
        $errorText = [string](Get-OptionalProperty -Object $result -Name 'error')
        if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = "counter exited $($captured.exit_code)" }
        throw "Could not count EndNote references: $errorText"
    }
    return [int](Get-OptionalProperty -Object $result -Name 'count')
}

function ConvertTo-ComparableVersion {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $script:ZeroVersion }
    $match = [regex]::Match($Text, '\d+(\.\d+)*')
    if (-not $match.Success) { return $script:ZeroVersion }
    $value = $match.Value
    if ($value -notmatch '\.') { $value = $value + '.0' }
    try { return [version]$value }
    catch {
        Write-Log -Level WARN -Message "Could not interpret '$Text' as a version: $($_.Exception.Message)"
        return $script:ZeroVersion
    }
}

function ConvertTo-XmlText {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function ConvertFrom-XmlText {
    param([Parameter(Mandatory)][string]$Value)
    return $Value.Replace('&lt;', '<').Replace('&gt;', '>').Replace('&amp;', '&')
}

# ---------------------------------------------------------------------------------------------
# EndNote installation discovery. The major version is never hardcoded.
# ---------------------------------------------------------------------------------------------

function Resolve-EndNoteDirectory {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $candidate = $Path.Trim().Trim('"')
    if ($candidate.Length -eq 0) { return $null }
    try {
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            if (Test-Path -LiteralPath (Join-Path $candidate $script:EndNoteExecutableName) -PathType Leaf) {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
            return $null
        }
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            if ([IO.Path]::GetFileName($candidate) -ieq $script:EndNoteExecutableName) {
                return (Resolve-Path -LiteralPath (Split-Path -Parent $candidate)).Path
            }
        }
    }
    catch {
        Write-Log -Level WARN -Message "Could not inspect candidate EndNote path '$candidate': $($_.Exception.Message)"
    }
    return $null
}

function New-EndNoteCandidate {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [AllowNull()][string]$VersionHint
    )
    $executable = Join-Path $Directory $script:EndNoteExecutableName
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) { return $null }
    $productVersion = $null
    try { $productVersion = [string](Get-Item -LiteralPath $executable).VersionInfo.ProductVersion }
    catch { Write-Log -Level WARN -Message "Could not read the product version of '$executable': $($_.Exception.Message)" }
    $version = ConvertTo-ComparableVersion -Text $VersionHint
    if ($version -eq $script:ZeroVersion) { $version = ConvertTo-ComparableVersion -Text $productVersion }
    $stylesDirectory = Join-Path $Directory $script:EndNoteStylesFolderName
    if (-not (Test-Path -LiteralPath $stylesDirectory -PathType Container)) { $stylesDirectory = $null }
    return [ordered]@{
        directory        = $Directory
        executable       = $executable
        product_version  = $productVersion
        version          = $version
        styles_directory = $stylesDirectory
    }
}

function Get-EndNoteInstallation {
    if ($null -ne $script:EndNoteInstallationCache) { return $script:EndNoteInstallationCache }

    $candidates = New-Object System.Collections.Generic.List[object]
    $seenDirectories = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)

    $addCandidate = {
        param($Directory, $VersionHint)
        if ([string]::IsNullOrWhiteSpace($Directory)) { return }
        if (-not $seenDirectories.Add($Directory)) { return }
        $candidate = New-EndNoteCandidate -Directory $Directory -VersionHint $VersionHint
        if ($null -ne $candidate) { $candidates.Add($candidate) }
    }

    # 1. Product registry keys. Value names differ between EndNote releases, so every string
    #    value of every version subkey is tested for an install directory instead of guessing.
    foreach ($root in $script:EndNoteRegistryRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $subKeys = @()
        try { $subKeys = @(Get-ChildItem -LiteralPath $root -ErrorAction Stop) }
        catch {
            Write-Log -Level WARN -Message "Could not enumerate '$root': $($_.Exception.Message)"
            continue
        }
        foreach ($subKey in $subKeys) {
            $properties = $null
            try { $properties = Get-ItemProperty -LiteralPath $subKey.PSPath -ErrorAction Stop }
            catch {
                Write-Log -Level WARN -Message "Could not read registry values of '$($subKey.Name)': $($_.Exception.Message)"
                continue
            }
            foreach ($property in $properties.PSObject.Properties) {
                if ($property.Value -isnot [string]) { continue }
                $directory = Resolve-EndNoteDirectory -Path $property.Value
                & $addCandidate $directory $subKey.PSChildName
            }
        }
    }

    # 2. App Paths registration.
    foreach ($appPathKey in $script:EndNoteAppPathsKeys) {
        if (-not (Test-Path -LiteralPath $appPathKey)) { continue }
        try {
            $appPath = Get-ItemProperty -LiteralPath $appPathKey -ErrorAction Stop
            $defaultValue = Get-OptionalProperty -Object $appPath -Name '(default)'
            $directory = Resolve-EndNoteDirectory -Path ([string]$defaultValue)
            & $addCandidate $directory $null
        }
        catch {
            Write-Log -Level WARN -Message "Could not read '$appPathKey': $($_.Exception.Message)"
        }
    }

    # 3. Program Files scan, 32-bit location first because EndNote ships as a 32-bit product.
    foreach ($programFiles in @(${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
        if ([string]::IsNullOrWhiteSpace($programFiles)) { continue }
        if (-not (Test-Path -LiteralPath $programFiles -PathType Container)) { continue }
        $directories = @()
        try {
            $directories = @(Get-ChildItem -LiteralPath $programFiles -Directory -Filter $script:EndNoteDirectoryFilter -ErrorAction Stop)
        }
        catch {
            Write-Log -Level WARN -Message "Could not enumerate '$programFiles': $($_.Exception.Message)"
            continue
        }
        foreach ($directory in $directories) {
            & $addCandidate $directory.FullName $directory.Name
        }
    }

    if ($candidates.Count -eq 0) {
        Write-Log -Level WARN -Message 'No EndNote installation was discovered in the registry or in Program Files.'
        return $null
    }

    $ordered = @($candidates | Sort-Object -Property @{ Expression = { $_.version } } -Descending)
    $script:EndNoteInstallationCache = $ordered[0]
    Write-Log -Level INFO -Message "Using EndNote installation: $($script:EndNoteInstallationCache.executable) (version $($script:EndNoteInstallationCache.product_version))"
    return $script:EndNoteInstallationCache
}

function Get-CwywTemplatePath {
    if ($null -ne $script:CwywTemplateCache) { return $script:CwywTemplateCache }
    $best = $null
    $bestVersion = $null
    foreach ($commonFiles in @($env:CommonProgramFiles, ${env:CommonProgramFiles(x86)})) {
        if ([string]::IsNullOrWhiteSpace($commonFiles)) { continue }
        $root = Join-Path $commonFiles $script:CwywRelativePath
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        $directories = @()
        try { $directories = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction Stop) }
        catch {
            Write-Log -Level WARN -Message "Could not enumerate '$root': $($_.Exception.Message)"
            continue
        }
        foreach ($directory in $directories) {
            $file = Join-Path $directory.FullName $script:CwywTemplateFileName
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }
            $version = ConvertTo-ComparableVersion -Text $directory.Name
            if ($null -eq $bestVersion -or $version -gt $bestVersion) {
                $best = $file
                $bestVersion = $version
            }
        }
    }
    if ($null -eq $best) {
        Write-Log -Level WARN -Message 'The EndNote CWYW Word template was not found under Common Files\ResearchSoft\Cwyw.'
    }
    $script:CwywTemplateCache = $best
    return $best
}

function Get-LibraryDataDirectory {
    param([Parameter(Mandatory)][string]$LibraryFile)
    $parent = Split-Path -Parent $LibraryFile
    $stem = [IO.Path]::GetFileNameWithoutExtension($LibraryFile)
    return (Join-Path $parent ($stem + $script:LibraryDataDirectorySuffix))
}

function Assert-EndNoteLibrary {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'An EndNote library path is required (manifest library.path or -LibraryPath).'
    }
    $resolved = Resolve-ExistingFile -Path $Path -Label 'library.path'
    $dataDirectory = Get-LibraryDataDirectory -LibraryFile $resolved
    if (-not (Test-Path -LiteralPath $dataDirectory -PathType Container)) {
        throw "The EndNote library data directory is missing: $dataDirectory. A .enl file without its sibling .Data directory holds no references."
    }
    return [ordered]@{ path = $resolved; data_directory = $dataDirectory }
}

# ---------------------------------------------------------------------------------------------
# Output styles.
# ---------------------------------------------------------------------------------------------

function Get-InstalledStyles {
    if ($null -ne $script:InstalledStylesCache) { return $script:InstalledStylesCache }

    $directories = New-Object System.Collections.Generic.List[string]
    $installation = Get-EndNoteInstallation
    if ($null -ne $installation -and -not [string]::IsNullOrWhiteSpace($installation.styles_directory)) {
        $directories.Add($installation.styles_directory)
    }
    $documents = [Environment]::GetFolderPath('MyDocuments')
    if (-not [string]::IsNullOrWhiteSpace($documents)) {
        $personal = Join-Path $documents $script:PersonalStylesRelativePath
        if (Test-Path -LiteralPath $personal -PathType Container) { $directories.Add($personal) }
    }

    $names = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    foreach ($directory in $directories) {
        try {
            foreach ($file in Get-ChildItem -LiteralPath $directory -File -Filter $script:EndNoteStyleFilter -ErrorAction Stop) {
                if ([IO.Path]::GetExtension($file.Name) -ine $script:EndNoteStyleExtension) { continue }
                [void]$names.Add([IO.Path]::GetFileNameWithoutExtension($file.Name))
            }
        }
        catch {
            Write-Log -Level WARN -Message "Could not enumerate styles in '$directory': $($_.Exception.Message)"
        }
    }

    $sorted = @($names | Sort-Object)
    $script:InstalledStylesCache = [ordered]@{
        directories = @($directories)
        names       = $sorted
    }
    return $script:InstalledStylesCache
}

function Resolve-StyleName {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'An output style name is required.' }
    $styles = Get-InstalledStyles
    if (@($styles.names).Count -eq 0) {
        throw 'No EndNote output styles (.ens) were found. Install EndNote or check the styles directory before rendering.'
    }
    foreach ($candidate in $styles.names) {
        if ($candidate -ieq $Name) { return $candidate }
    }
    $escaped = [System.Management.Automation.WildcardPattern]::Escape($Name)
    $near = @($styles.names | Where-Object { $_ -like "*$escaped*" } | Select-Object -First $script:MaxStyleSuggestions)
    $suggestion = 'none'
    if ($near.Count -gt 0) { $suggestion = ($near -join ', ') }
    throw "Output style '$Name' is not installed. An unknown style makes EndNote open a modal dialog and hang the run, so this is a hard failure. Nearest installed matches: $suggestion. Use the list-styles action for the full list."
}

# ---------------------------------------------------------------------------------------------
# Word lifecycle.
# ---------------------------------------------------------------------------------------------

function Start-Word {
    param([switch]$ReadOnly, [string]$DocumentPath)
    $before = @(Get-Process $script:WordProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $script:Word = New-Object -ComObject Word.Application
    $script:Word.Visible = $false
    $script:Word.DisplayAlerts = $script:WordAlertsNone
    if ([string]::IsNullOrWhiteSpace($DocumentPath)) {
        $script:Document = $script:Word.Documents.Add()
    }
    else {
        $script:Document = $script:Word.Documents.Open($DocumentPath, $false, [bool]$ReadOnly)
    }
    $newProcesses = @(Get-Process $script:WordProcessName -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($newProcesses.Count -eq 1) { Write-Log -Level INFO -Message "WORD_PID=$($newProcesses[0].Id)" }
    elseif ($newProcesses.Count -gt 1) { Write-Log -Level WARN -Message "Multiple new Word PIDs observed: $($newProcesses.Id -join ', ')" }
    $script:Selection = $script:Word.Selection
}

function Stop-Word {
    param([switch]$Save)
    if ($null -ne $script:Document) {
        $saveOption = $script:WdDoNotSaveChanges
        if ($Save) { $saveOption = $script:WdSaveChanges }
        try { $script:Document.Close($saveOption) }
        catch { Write-Log -Level WARN -Message "Could not close Word document cleanly: $($_.Exception.Message)" }
    }
    if ($null -ne $script:Word) {
        try { $script:Word.Quit() }
        catch { Write-Log -Level WARN -Message "Could not quit Word cleanly: $($_.Exception.Message)" }
    }
    Release-ComObject $script:Selection
    Release-ComObject $script:Document
    Release-ComObject $script:Word
    $script:Selection = $null
    $script:Document = $null
    $script:Word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function New-WorkingCopy {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$Destination
    )
    # The input document is user data and is never opened for writing. Word always edits a
    # private copy that lives beside the destination so the final Move-Item stays on one volume.
    foreach ($stalePath in @(Remove-EndNoteStaleDocumentFiles -Destination $Destination)) {
        $script:StaleArtifactsRemoved.Add($stalePath) | Out-Null
        Write-Log -Level INFO -Message "Removed stale generated document: $stalePath"
    }
    $directory = Split-Path -Parent $Destination
    $name = '.' + [IO.Path]::GetFileNameWithoutExtension($Destination) + '.' + [Guid]::NewGuid().ToString('N') + $script:WorkingCopySuffix
    $path = Join-Path $directory $name
    $script:WorkingCopyPath = $path
    $script:WorkingCopyDestination = $Destination
    Copy-Item -LiteralPath $SourcePath -Destination $path -Force
    Write-Log -Level INFO -Message "Created working copy: $path"
    return $path
}

function Remove-WorkingCopy {
    if ([string]::IsNullOrWhiteSpace($script:WorkingCopyPath)) { return }
    $path = $script:WorkingCopyPath
    $destination = $script:WorkingCopyDestination
    try {
        if (Remove-EndNoteGeneratedDocumentFile -Path $path -Destination $destination) {
            $script:CleanupRemoved.Add($path) | Out-Null
        }
    }
    catch {
        $message = "Could not delete and verify the working copy '$path': $($_.Exception.Message)"
        $script:CleanupFailures.Add($message) | Out-Null
        Write-Log -Level ERROR -Message $message
    }
    finally {
        $script:WorkingCopyPath = $null
        $script:WorkingCopyDestination = $null
    }
}

function Remove-PendingTemporarySave {
    if ([string]::IsNullOrWhiteSpace($script:TemporarySavePath)) { return }
    $path = $script:TemporarySavePath
    $destination = $script:TemporarySaveDestination
    try {
        if (Remove-EndNoteGeneratedDocumentFile -Path $path -Destination $destination) {
            $script:CleanupRemoved.Add($path) | Out-Null
        }
    }
    catch {
        $message = "Could not delete and verify the temporary save '$path': $($_.Exception.Message)"
        $script:CleanupFailures.Add($message) | Out-Null
        Write-Log -Level ERROR -Message $message
    }
    finally {
        $script:TemporarySavePath = $null
        $script:TemporarySaveDestination = $null
    }
}

function Assert-EndNoteCleanupSucceeded {
    if ($script:CleanupFailures.Count -gt 0) {
        throw ('Cleanup verification failed: ' + ($script:CleanupFailures -join '; '))
    }
}

function Get-EndNoteCleanupReport {
    return [ordered]@{
        ok            = $script:CleanupFailures.Count -eq 0
        removed       = @($script:CleanupRemoved)
        stale_removed = @($script:StaleArtifactsRemoved)
    }
}

function Save-DocumentAtomically {
    param([Parameter(Mandatory)][string]$Destination)
    $temporary = Join-Path (Split-Path -Parent $Destination) ('.' + [IO.Path]::GetFileNameWithoutExtension($Destination) + '.' + [Guid]::NewGuid().ToString('N') + $script:TemporaryDocumentSuffix)
    $script:TemporarySavePath = $temporary
    $script:TemporarySaveDestination = $Destination
    try {
        Write-Log -Level INFO -Message "Saving temporary DOCX: $temporary"
        $script:Document.SaveAs2($temporary, $script:WdFormatXmlDocument)
        Write-Log -Level INFO -Message 'Word SaveAs2 returned.'
        Stop-Word
        Move-Item -LiteralPath $temporary -Destination $Destination -Force
        Write-Log -Level INFO -Message "Moved completed DOCX to: $Destination"
    }
    finally {
        Remove-PendingTemporarySave
    }
}

# ---------------------------------------------------------------------------------------------
# CWYW entry points. Route B (Application.Run) only; COMAddIns(...).Object returns $null.
# ---------------------------------------------------------------------------------------------

function Assert-CwywTemplateLoaded {
    if ($null -eq $script:Word) { throw 'Word is not running; the CWYW template cannot be checked.' }
    $loaded = $false
    foreach ($template in @($script:Word.Templates)) {
        if ($template.Name -like $script:CwywTemplateNamePattern) { $loaded = $true }
        Release-ComObject $template
    }
    if (-not $loaded) {
        throw "The '$script:CwywTemplateFileName' global template is not loaded in Word. EndNote loads it automatically once a document is open; if it is missing, repair the EndNote CWYW installation. Do not call Word.AddIns.Add, which double-loads the template."
    }
}

function Invoke-CwywMacro {
    param([Parameter(Mandatory)][string]$Name)
    if ($Name -match $script:InteractiveMacroPattern) {
        throw "Refusing to call the interactive CWYW macro '$Name'. Macros whose name ends in Dlg open a modal dialog that blocks Application.Run forever."
    }
    if ($script:AllowedCwywMacros -notcontains $Name) {
        throw "Refusing to call unverified CWYW macro '$Name'. Only measured, non-interactive macros are allowed."
    }
    if ($null -eq $script:Word) { throw "Word is not running; cannot call '$Name'." }
    Assert-CwywTemplateLoaded
    Write-Log -Level INFO -Message "Calling CWYW macro '$Name'."
    $null = $script:Word.Run($Name)
    Write-Log -Level INFO -Message "CWYW macro '$Name' returned."
}

function Invoke-CwywFormat {
    Assert-NoUnresolvedCitationMarkers
    Invoke-CwywMacro -Name $script:CwywFormatMacro
}

function Get-EndNoteProcessIds {
    # Snapshot taken before anything is launched. CWYW auto-starts EndNote by itself during
    # ENFormatBibliography when no instance is running, so a new PID can appear without the
    # bridge ever calling Start-Process.
    $ids = New-Object System.Collections.Generic.HashSet[int]
    try {
        foreach ($process in @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue)) {
            [void]$ids.Add([int]$process.Id)
        }
    }
    catch {
        Write-Log -Level WARN -Message "Could not enumerate EndNote processes: $($_.Exception.Message)"
    }
    # Comma wrapper: PowerShell unrolls any IEnumerable written to the output stream, so a bare
    # "return $ids" would hand back loose integers (or $null when nothing is running).
    return , $ids
}

function Test-EndNoteAutostarted {
    param([Parameter(Mandatory)]$KnownProcessIds)
    try {
        foreach ($process in @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue)) {
            if (-not $KnownProcessIds.Contains([int]$process.Id)) { return $true }
        }
    }
    catch {
        Write-Log -Level WARN -Message "Could not compare EndNote processes: $($_.Exception.Message)"
    }
    return $false
}

function Start-EndNoteWithLibrary {
    param([Parameter(Mandatory)][string]$LibraryFile)
    $running = @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($running.Count -gt 0) {
        Write-Log -Level INFO -Message "EndNote is already running (PID $($running[0].Id)). Not starting a second instance and not switching its open library."
        return [ordered]@{ preexisting = $true; process_id = $running[0].Id }
    }

    $installation = Get-EndNoteInstallation
    if ($null -eq $installation) {
        throw 'EndNote is not installed or could not be discovered, so the CWYW formatter has no library server to talk to.'
    }

    Write-Log -Level INFO -Message "Starting EndNote with library: $LibraryFile"
    $argument = '"' + $LibraryFile + '"'
    $started = $null
    try {
        $started = Start-Process -FilePath $installation.executable -ArgumentList $argument -PassThru -WindowStyle Hidden
    }
    catch { throw "Could not start EndNote from '$($installation.executable)': $($_.Exception.Message)" }

    $deadline = [DateTime]::UtcNow.AddMilliseconds($script:EndNoteReadyTimeoutMilliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $ready = @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($ready.Count -gt 0) {
            Write-Log -Level INFO -Message "ENDNOTE_PID=$($ready[0].Id)"
            return [ordered]@{ preexisting = $false; process_id = $ready[0].Id }
        }
        Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
    }

    $startedId = 'unknown'
    if ($null -ne $started) { $startedId = [string]$started.Id }
    throw "EndNote (PID $startedId) did not open a main window within $($script:EndNoteReadyTimeoutMilliseconds) ms. Raise ENDNOTE_READY_TIMEOUT_MS or open the library manually first."
}

function Get-EndNoteMostRecentLibraryPath {
    try {
        $recent = Get-ItemProperty -LiteralPath $script:EndNoteRecentLibrariesRegistryPath -ErrorAction Stop
        $value = [string](Get-OptionalProperty -Object $recent -Name $script:EndNoteMostRecentLibraryValueName)
        if ([string]::IsNullOrWhiteSpace($value)) { return $null }
        return [IO.Path]::GetFullPath($value)
    }
    catch {
        return $null
    }
}

function Resolve-EndNoteImportLibrary {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [AllowNull()][string]$EmptyTemplatePath
    )
    $fullTarget = [IO.Path]::GetFullPath($TargetPath)
    if ([IO.Path]::GetExtension($fullTarget) -ine '.enl') {
        throw "LibraryPath must use the .enl extension: $fullTarget"
    }
    if (Test-Path -LiteralPath $fullTarget -PathType Leaf) {
        $library = Assert-EndNoteLibrary -Path $fullTarget
        return [ordered]@{ library = $library; created = $false; template_path = $null }
    }

    $targetData = Get-LibraryDataDirectory -LibraryFile $fullTarget
    if (Test-Path -LiteralPath $targetData) {
        throw "Cannot create the library because its sibling data path already exists: $targetData"
    }
    if ([string]::IsNullOrWhiteSpace($EmptyTemplatePath)) {
        throw 'LibraryPath does not exist. Supply EmptyLibraryTemplatePath pointing to a closed, empty EndNote library to create a new library before import.'
    }

    $template = Assert-EndNoteLibrary -Path $EmptyTemplatePath
    $runningEndNote = @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue)
    if ($runningEndNote.Count -gt 0) {
        throw 'Creating a new library from EmptyLibraryTemplatePath requires EndNote to be fully closed, so no background library or SQLite/WAL file can change during the pair copy. Close EndNote and retry.'
    }
    $templateCount = Get-EndNoteRecordCount -LibraryFile $template.path -IncludeTrash
    if ($templateCount -ne 0) {
        throw "EmptyLibraryTemplatePath must contain zero stored references, including trash; found $templateCount in $($template.path)."
    }
    $parent = Split-Path -Parent $fullTarget
    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw "LibraryPath must be an absolute path with a parent directory: $fullTarget"
    }
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    Write-Log -Level INFO -Message "Creating a new EndNote library from the verified empty template: $($template.path)"
    try {
        Copy-Item -LiteralPath $template.path -Destination $fullTarget -ErrorAction Stop
        Copy-Item -LiteralPath $template.data_directory -Destination $targetData -Recurse -ErrorAction Stop
    }
    catch {
        $copyError = $_.Exception.Message
        try {
            if (Test-Path -LiteralPath $targetData) {
                Remove-Item -LiteralPath $targetData -Recurse -Force -ErrorAction Stop
            }
            if (Test-Path -LiteralPath $fullTarget) {
                Remove-Item -LiteralPath $fullTarget -Force -ErrorAction Stop
            }
        }
        catch {
            throw "Could not copy the empty EndNote library template, and cleanup of the newly created partial target also failed: $copyError | $($_.Exception.Message)"
        }
        throw "Could not copy the empty EndNote library template. The partial target was removed. Close the template library in EndNote and retry: $copyError"
    }
    $created = Assert-EndNoteLibrary -Path $fullTarget
    if ((Get-EndNoteRecordCount -LibraryFile $created.path) -ne 0) {
        throw "The newly copied EndNote library is not empty: $($created.path)"
    }
    return [ordered]@{ library = $created; created = $true; template_path = $template.path }
}

function Open-EndNoteLibraryForImport {
    param(
        [Parameter(Mandatory)][string]$LibraryFile,
        [Parameter(Mandatory)]$Installation
    )
    $fileName = [IO.Path]::GetFileName($LibraryFile)
    $fullLibrary = [IO.Path]::GetFullPath($LibraryFile)
    $preexisting = $false
    $recentBefore = Get-EndNoteMostRecentLibraryPath
    if (-not [string]::IsNullOrWhiteSpace($recentBefore) -and
        [string]::Equals($recentBefore, $fullLibrary, [StringComparison]::OrdinalIgnoreCase)) {
        foreach ($process in @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue)) {
            if ($process.MainWindowHandle -ne 0 -and $process.MainWindowTitle -like "* - $fileName") {
                $preexisting = $true
                break
            }
        }
    }

    # Unlike render, import is an explicit request to mutate this exact library. Opening the
    # supplied path is therefore required even when another library is already open.
    Write-Log -Level INFO -Message "Opening the import target library: $LibraryFile"
    Start-Process -FilePath $Installation.executable -ArgumentList ('"' + $LibraryFile + '"') -WindowStyle Hidden | Out-Null
    $deadline = [DateTime]::UtcNow.AddMilliseconds($script:EndNoteReadyTimeoutMilliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        $recentLibrary = Get-EndNoteMostRecentLibraryPath
        foreach ($process in @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue)) {
            if ($process.MainWindowHandle -ne 0 -and $process.MainWindowTitle -like "* - $fileName" -and
                -not [string]::IsNullOrWhiteSpace($recentLibrary) -and
                [string]::Equals($recentLibrary, $fullLibrary, [StringComparison]::OrdinalIgnoreCase)) {
                return [ordered]@{
                    preexisting = $preexisting
                    process_id = $process.Id
                    title = $process.MainWindowTitle
                    confirmed_library_path = $recentLibrary
                }
            }
        }
        Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
    }
    throw "EndNote did not confirm the exact requested import target '$fullLibrary' within $($script:EndNoteReadyTimeoutMilliseconds) ms."
}

# ---------------------------------------------------------------------------------------------
# Document variables: EN.Layout and EN.Libraries.
# ---------------------------------------------------------------------------------------------

function Get-DocumentVariable {
    param([Parameter(Mandatory)][string]$Name)
    if ($null -eq $script:Document) { throw "No Word document is open; cannot read the '$Name' document variable." }
    $result = $null
    for ($index = 1; $index -le $script:Document.Variables.Count; $index++) {
        $variable = $script:Document.Variables.Item($index)
        try {
            if ([string]$variable.Name -eq $Name) { $result = [string]$variable.Value }
        }
        finally { Release-ComObject $variable }
        if ($null -ne $result) { break }
    }
    return $result
}

function Set-DocumentVariable {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Value
    )
    if ($null -eq $script:Document) { throw "No Word document is open; cannot write the '$Name' document variable." }
    $updated = $false
    for ($index = 1; $index -le $script:Document.Variables.Count; $index++) {
        $variable = $script:Document.Variables.Item($index)
        try {
            if ([string]$variable.Name -eq $Name) {
                $variable.Value = $Value
                $updated = $true
            }
        }
        finally { Release-ComObject $variable }
        if ($updated) { break }
    }
    if (-not $updated) {
        $created = $script:Document.Variables.Add($Name, $Value)
        Release-ComObject $created
    }
}

function Get-EndNoteLayout {
    $value = Get-DocumentVariable -Name $script:LayoutVariableName
    if ($null -eq $value) { return $null }
    $style = $null
    $match = [regex]::Match($value, $script:LayoutStylePattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    if ($match.Success) { $style = ConvertFrom-XmlText -Value $match.Groups[1].Value }
    return [ordered]@{ value = $value; style = $style }
}

function Set-EndNoteLayout {
    param([Parameter(Mandatory)][string]$StyleName)
    $escaped = ConvertTo-XmlText -Value $StyleName
    $element = '<Style>' + $escaped + '</Style>'
    $layout = Get-EndNoteLayout
    $newValue = $null
    if ($null -eq $layout) {
        # The document has never been formatted by EndNote, so no EN.Layout exists yet.
        $newValue = $script:DefaultLayoutTemplate.Replace($script:LayoutStylePlaceholder, $escaped)
        Write-Log -Level INFO -Message "Created the $($script:LayoutVariableName) document variable with style '$StyleName'."
    }
    else {
        $current = [string]$layout.value
        $match = [regex]::Match($current, $script:LayoutStylePattern, [Text.RegularExpressions.RegexOptions]::Singleline)
        if ($match.Success) {
            # Plain string surgery, not Regex.Replace, so that a style name containing $ or \
            # cannot be reinterpreted as a substitution pattern.
            $newValue = $current.Substring(0, $match.Index) + $element + $current.Substring($match.Index + $match.Length)
        }
        elseif ($current.StartsWith($script:LayoutRootOpenTag)) {
            $newValue = $current.Insert($script:LayoutRootOpenTag.Length, $element)
        }
        else {
            throw "The $($script:LayoutVariableName) document variable does not look like EndNote layout XML and was left untouched: $current"
        }
        Write-Log -Level INFO -Message "Set the $($script:LayoutVariableName) style to '$StyleName' (previous: '$($layout.style)')."
    }
    Set-DocumentVariable -Name $script:LayoutVariableName -Value $newValue
    $previous = $null
    if ($null -ne $layout) { $previous = $layout.style }
    return [ordered]@{ previous_style = $previous; style = $StyleName }
}

function Get-EndNoteLibraries {
    param([Parameter(Mandatory)]$Warnings)
    $libraries = New-Object System.Collections.Generic.List[object]
    $value = Get-DocumentVariable -Name $script:LibrariesVariableName
    # Every return is comma-wrapped: PowerShell unrolls an IEnumerable written to the output
    # stream, so a bare "return $libraries" hands back $null for an empty list (which then makes
    # the caller's .Count throw under Set-StrictMode) and a bare OrderedDictionary for a single
    # library, whose .Count is its key count rather than one.
    #
    # The list is also converted with ToArray() before it leaves this function. Both PowerShell
    # 7.6.3 and Windows PowerShell 5.1 on this platform throw "Argument types do not match" for
    # the array subexpression @($x) when $x is a List[object], which silently broke the whole
    # validate action at its final result-building step. Handing callers a plain object[]
    # removes that trap instead of relying on every call site to avoid @().
    if ($null -eq $value) { return , $libraries.ToArray() }
    $document = $null
    try { $document = [xml]$value }
    catch {
        $Warnings.Add("The $($script:LibrariesVariableName) document variable is not well-formed XML: $($_.Exception.Message)")
        return , $libraries.ToArray()
    }
    foreach ($node in $document.SelectNodes('/Libraries/item')) {
        $name = ''
        $textNode = $node.SelectSingleNode('text()')
        if ($null -ne $textNode) { $name = [string]$textNode.Value }
        $recordIds = New-Object System.Collections.Generic.List[int]
        foreach ($recordNode in $node.SelectNodes('record-ids/item')) {
            $parsed = 0
            if ([int]::TryParse([string]$recordNode.InnerText, [ref]$parsed)) { $recordIds.Add($parsed) }
            else { $Warnings.Add("Ignoring a non-numeric record id in $($script:LibrariesVariableName): $($recordNode.InnerText)") }
        }
        $libraries.Add([ordered]@{
                db_id      = [string]$node.GetAttribute('db-id')
                name       = $name.Trim()
                record_ids = $recordIds.ToArray()
            })
    }
    return , $libraries.ToArray()
}

# ---------------------------------------------------------------------------------------------
# Manifest handling.
# ---------------------------------------------------------------------------------------------

function Read-Manifest {
    param([Parameter(Mandatory)][string]$Path)
    $resolved = Resolve-ExistingFile -Path $Path -Label 'ManifestPath'
    $manifest = $null
    # -Encoding UTF8: Windows PowerShell 5.1 would otherwise read a BOM-less UTF-8 manifest with
    # the system ANSI codepage, which mangles non-ASCII author names into invalid JSON.
    try { $manifest = Get-Content -Raw -Encoding UTF8 -LiteralPath $resolved | ConvertFrom-Json }
    catch { throw "ManifestPath is not valid JSON ($resolved): $($_.Exception.Message)" }
    if ($null -eq $manifest) { throw "ManifestPath is empty: $resolved" }
    $schemaVersion = Get-OptionalProperty -Object $manifest -Name 'schema_version'
    if ($schemaVersion -ne $script:ManifestSchemaVersion) {
        throw "Manifest schema_version must be $($script:ManifestSchemaVersion)."
    }
    return $manifest
}

function Assert-AllowedKeys {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string[]]$Allowed,
        [Parameter(Mandatory)][string]$Label
    )
    foreach ($property in $Object.PSObject.Properties) {
        if ($Allowed -notcontains $property.Name) {
            throw "$Label contains the unsupported key '$($property.Name)'. Allowed keys: $($Allowed -join ', ')."
        }
    }
}

function ConvertTo-ManifestInteger {
    param(
        [Parameter(Mandatory)][AllowNull()]$Value,
        [Parameter(Mandatory)][string]$Label
    )
    if ($null -eq $Value) { throw "$Label is required and must be a JSON integer." }
    if ($Value -is [string]) { throw "$Label must be a JSON number, not a quoted string." }
    if ($Value -is [bool]) { throw "$Label must be a JSON integer, not a boolean." }
    $text = [string]$Value
    $parsed = 0
    if (-not [int]::TryParse($text, [ref]$parsed)) { throw "$Label must be an integer: $text" }
    return $parsed
}

function Assert-CitationText {
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][string]$Label
    )
    if ($Value -match $script:ForbiddenCitationTextPattern) {
        throw "$Label must not contain any of { } # @ ; or a backslash. Those characters break the generated temporary citation or trigger an interactive EndNote dialog. See references/temporary-citations.md."
    }
}

function Get-CitationDisplay {
    param([Parameter(Mandatory)]$Citation)
    $display = [string](Get-OptionalProperty -Object $Citation -Name 'display')
    if ([string]::IsNullOrWhiteSpace($display)) { return $script:DefaultCitationDisplay }
    return $display.Trim().ToLowerInvariant()
}

function Assert-ManifestTarget {
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)][string]$CitationId
    )
    Assert-AllowedKeys -Object $Target -Allowed $script:AllowedTargetKeys -Label "A target of citation '$CitationId'"
    $recNumber = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $Target -Name 'rec_number') -Label "Citation '$CitationId' rec_number"
    if ($recNumber -le 0) { throw "Citation '$CitationId' rec_number must be greater than 0: $recNumber" }
    $author = [string](Get-OptionalProperty -Object $Target -Name 'author')
    if ([string]::IsNullOrWhiteSpace($author)) { throw "Citation '$CitationId' requires a non-empty author for every target." }
    Assert-CitationText -Value $author -Label "Citation '$CitationId' author"
    $null = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $Target -Name 'year') -Label "Citation '$CitationId' year"
    $doi = [string](Get-OptionalProperty -Object $Target -Name 'doi')
    if ([string]::IsNullOrWhiteSpace($doi)) { throw "Citation '$CitationId' requires a DOI for every target." }
    if ($doi -notmatch $script:DoiPattern) {
        throw "Citation '$CitationId' DOI must be the bare DOI without 'doi:' or a resolver URL: $doi"
    }
}

function Assert-CitationOptions {
    param(
        [Parameter(Mandatory)]$Citation,
        [Parameter(Mandatory)][string]$CitationId,
        [Parameter(Mandatory)][int]$TargetCount
    )
    $display = Get-CitationDisplay -Citation $Citation
    if ($script:AllowedCitationDisplays -notcontains $display) {
        throw "Citation '$CitationId' display must be 'default' or 'author-year': $display"
    }
    if ($display -eq $script:AuthorYearCitationDisplay -and $TargetCount -ne 1) {
        throw "Citation '$CitationId' uses display 'author-year', which requires exactly one target. Use default display for merged citations."
    }
    $pages = Get-OptionalProperty -Object $Citation -Name 'pages'
    $suffix = Get-OptionalProperty -Object $Citation -Name 'suffix'
    if ($null -ne $pages) {
        if ($pages -isnot [string]) { throw "Citation '$CitationId' pages must be a string." }
        if ([string]$pages -notmatch $script:CitedPagesPattern) {
            throw "Citation '$CitationId' pages must match $($script:CitedPagesPattern), for example 45 or 45-46: $pages"
        }
    }
    if ($null -ne $suffix) {
        if ($suffix -isnot [string]) { throw "Citation '$CitationId' suffix must be a string." }
        if ([string]::IsNullOrWhiteSpace([string]$suffix)) { throw "Citation '$CitationId' suffix must not be blank." }
        Assert-CitationText -Value ([string]$suffix) -Label "Citation '$CitationId' suffix"
    }
    if (($null -ne $pages -or $null -ne $suffix) -and $TargetCount -ne 1) {
        throw "Citation '$CitationId' supplies pages or suffix, which EndNote applies to the last target only. Split the merged citation so that pages and suffix belong to a citation with exactly one target."
    }
}

function Assert-ManifestCitation {
    param(
        [Parameter(Mandatory)]$Citation,
        [Parameter(Mandatory)]$Identifiers,
        [Parameter(Mandatory)]$Markers
    )
    $id = [string](Get-OptionalProperty -Object $Citation -Name 'id')
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'Every citation requires a non-empty id.' }
    if ($id -notmatch $script:CitationIdPattern) { throw "Citation id must match $($script:CitationIdPattern): $id" }
    if (-not $Identifiers.Add($id)) { throw "Duplicate citation id: $id" }
    foreach ($rejected in $script:RejectedCitationKeys) {
        if ($null -ne $Citation.PSObject.Properties[$rejected]) {
            throw "Citation '$id' uses '$rejected', which schema v1 rejects. EndNote's backslash prefix syntax opens a modal dialog and hangs the run. See references/temporary-citations.md."
        }
    }
    Assert-AllowedKeys -Object $Citation -Allowed $script:AllowedCitationKeys -Label "Citation '$id'"
    $marker = [string](Get-OptionalProperty -Object $Citation -Name 'marker')
    if ([string]::IsNullOrWhiteSpace($marker)) { throw "Citation '$id' requires a marker." }
    if (-not $Markers.Add($marker)) { throw "Duplicate marker: $marker" }
    $targets = @(Get-OptionalProperty -Object $Citation -Name 'targets')
    if ($targets.Count -eq 0) { throw "Citation '$id' requires a non-empty targets array." }
    foreach ($target in $targets) { Assert-ManifestTarget -Target $target -CitationId $id }
    Assert-CitationOptions -Citation $Citation -CitationId $id -TargetCount $targets.Count
}

function Assert-Manifest {
    param(
        [Parameter(Mandatory)]$Manifest,
        [string]$LibraryOverride,
        [switch]$SkipLibraryCheck
    )
    Assert-AllowedKeys -Object $Manifest -Allowed $script:AllowedManifestKeys -Label 'The manifest'

    $styleValue = Get-OptionalProperty -Object $Manifest -Name 'style'
    if ($null -ne $styleValue -and ($styleValue -isnot [string] -or [string]::IsNullOrWhiteSpace([string]$styleValue))) {
        throw 'Manifest style must be omitted or be a non-empty string naming an installed EndNote output style.'
    }
    if ($null -eq $styleValue) { $null = Get-DefaultOutputStyle }

    $citations = @(Get-OptionalProperty -Object $Manifest -Name 'citations')
    if ($citations.Count -eq 0) { throw 'Manifest citations must be a non-empty array.' }

    $bibliography = Get-OptionalProperty -Object $Manifest -Name 'bibliography'
    if ($null -ne $bibliography) {
        Assert-AllowedKeys -Object $bibliography -Allowed $script:AllowedBibliographyKeys -Label 'manifest bibliography'
        $expected = Get-OptionalProperty -Object $bibliography -Name 'expected'
        if ($null -ne $expected -and $expected -isnot [bool]) {
            throw 'Manifest bibliography.expected must be a boolean.'
        }
    }

    $ids = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::Ordinal)
    $markers = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::Ordinal)
    foreach ($citation in $citations) {
        Assert-ManifestCitation -Citation $citation -Identifiers $ids -Markers $markers
    }

    if ($SkipLibraryCheck) { return }

    $libraryValue = $LibraryOverride
    if ([string]::IsNullOrWhiteSpace($libraryValue)) {
        $library = Get-OptionalProperty -Object $Manifest -Name 'library'
        if ($null -eq $library) { throw 'The manifest requires a library object with a path.' }
        Assert-AllowedKeys -Object $library -Allowed $script:AllowedLibraryKeys -Label 'manifest library'
        $libraryValue = [string](Get-OptionalProperty -Object $library -Name 'path')
    }
    $null = Assert-EndNoteLibrary -Path $libraryValue
}

function Get-TemporaryCitationText {
    param([Parameter(Mandatory)]$Citation)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($target in @(Get-OptionalProperty -Object $Citation -Name 'targets')) {
        $author = [string](Get-OptionalProperty -Object $target -Name 'author')
        $year = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $target -Name 'year') -Label 'year'
        $recNumber = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $target -Name 'rec_number') -Label 'rec_number'
        $parts.Add($author.Trim() + ', ' + [string]$year + ' #' + [string]$recNumber)
    }
    $text = '{' + ($parts -join '; ')
    $pages = Get-OptionalProperty -Object $Citation -Name 'pages'
    if ($null -ne $pages) { $text = $text + '@' + ([string]$pages).Trim() }
    $suffix = Get-OptionalProperty -Object $Citation -Name 'suffix'
    if ($null -ne $suffix) { $text = $text + [string]$suffix }
    return $text + '}'
}

function Get-CitationMarkerPlans {
    param([Parameter(Mandatory)][object[]]$Citations)
    $plans = New-Object System.Collections.Generic.List[object]
    foreach ($citation in $Citations) {
        $marker = [string](Get-OptionalProperty -Object $citation -Name 'marker')
        $range = Find-UniqueMarkerRange -Marker $marker
        try {
            # Must be a PSCustomObject: Sort-Object -Property cannot see the keys of an
            # [ordered] dictionary, so sorting plans by 'start' silently returns them in
            # manifest order. Replacing forwards then shifts every later offset and the
            # remaining markers get written over the wrong text.
            $plans.Add([pscustomobject][ordered]@{
                    citation = $citation
                    start    = [int]$range.Start
                    end      = [int]$range.End
                    text     = Get-TemporaryCitationText -Citation $citation
                })
        }
        finally { Release-ComObject $range }
    }
    return @($plans | Sort-Object -Property start)
}

function Set-CitationMarkers {
    param([Parameter(Mandatory)][object[]]$Plans)
    # Replace from the end of the document so every earlier saved range stays valid.
    foreach ($plan in @($Plans | Sort-Object -Property start -Descending)) {
        $range = $script:Document.Range([int]$plan.start, [int]$plan.end)
        try { $range.Text = [string]$plan.text }
        finally { Release-ComObject $range }
    }
}

function Assert-NoUnresolvedCitationMarkers {
    $content = $script:Document.Content
    try { $text = [string]$content.Text }
    finally { Release-ComObject $content }
    if ($text.IndexOf($script:CitationMarkerPrefix, [StringComparison]::Ordinal) -ge 0) {
        throw "Unresolved citation marker text remains after manifest replacement ($($script:CitationMarkerPrefix)...). Refusing to call CWYW because EndNote treats braces as temporary citations and may open Select Matching Reference. Add every document marker to the manifest, or remove unselected markers from a scratch copy before a subset test."
    }
}

function Get-CitationFieldByOrdinal {
    param([Parameter(Mandatory)][int]$Ordinal)
    $seen = 0
    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        $keep = $false
        try {
            $code = [string]$field.Code.Text
            if ($code -match $script:CitationDataFieldPattern) { continue }
            if ($code -match $script:CitationFieldPattern) {
                if ($seen -eq $Ordinal) {
                    $keep = $true
                    return $field
                }
                $seen++
            }
        }
        finally {
            if (-not $keep) { Release-ComObject $field }
        }
    }
    throw "Could not find native EndNote citation field ordinal $Ordinal."
}

function Set-CitationDisplayFormats {
    param(
        [Parameter(Mandatory)][object[]]$Plans,
        [Parameter(Mandatory)][string]$StyleName
    )
    $authorYearCount = 0
    for ($ordinal = 0; $ordinal -lt $Plans.Count; $ordinal++) {
        $citation = $Plans[$ordinal].citation
        if ((Get-CitationDisplay -Citation $citation) -ne $script:AuthorYearCitationDisplay) {
            continue
        }
        $field = Get-CitationFieldByOrdinal -Ordinal $ordinal
        $result = $null
        try {
            $result = $field.Result
            $result.Select()
            Invoke-CwywMacro -Name $script:CwywAuthorYearMacro
        }
        finally {
            Release-ComObject $result
            Release-ComObject $field
        }

        $updatedField = Get-CitationFieldByOrdinal -Ordinal $ordinal
        $updatedResult = $null
        try {
            $updatedResult = $updatedField.Result
            $displayText = [string]$updatedResult.Text
            $target = @(Get-OptionalProperty -Object $citation -Name 'targets')[0]
            $author = [string](Get-OptionalProperty -Object $target -Name 'author')
            if ($displayText.IndexOf($author, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                throw "Citation display conversion failed: '$displayText' does not contain author '$author'."
            }
            if ($StyleName -ieq 'IEEE' -and $displayText -notmatch '\[\s*\d+') {
                throw "IEEE author-year citation '$displayText' does not contain a bracketed bibliography number."
            }
        }
        finally {
            Release-ComObject $updatedResult
            Release-ComObject $updatedField
        }
        $authorYearCount++
    }
    return $authorYearCount
}

# ---------------------------------------------------------------------------------------------
# Document inspection.
# ---------------------------------------------------------------------------------------------

function Get-FieldResultText {
    param(
        [Parameter(Mandatory)]$Field,
        [Parameter(Mandatory)][string]$Label
    )
    try { return [string]$Field.Result.Text }
    catch {
        Write-Log -Level WARN -Message "Could not read the result of $Label`: $($_.Exception.Message)"
        return ''
    }
}

function Find-UniqueMarkerRange {
    param([Parameter(Mandatory)][string]$Marker)
    $search = $script:Document.Content.Duplicate
    # Local name deliberately differs from the automatic $Matches variable.
    $hits = @()
    try {
        while ($search.Start -lt $search.End) {
            $find = $search.Find
            $find.ClearFormatting()
            $find.Text = $Marker
            $find.Forward = $true
            $find.Wrap = $script:WdFindStop
            $find.MatchWildcards = $false
            if (-not $find.Execute()) { break }
            $hits += [ordered]@{ start = $search.Start; end = $search.End }
            $next = $search.End
            $search.SetRange($next, $script:Document.Content.End)
            Release-ComObject $find
        }
    }
    finally { Release-ComObject $search }
    if ($hits.Count -ne 1) {
        throw "Marker must occur exactly once; found $($hits.Count): $Marker"
    }
    return $script:Document.Range($hits[0].start, $hits[0].end)
}

function Get-CitationFields {
    if ($null -eq $script:Document) { throw 'No Word document is open; cannot enumerate fields.' }
    $citationCodes = New-Object System.Collections.Generic.List[string]
    $citationResults = New-Object System.Collections.Generic.List[string]
    $citationDataCodes = New-Object System.Collections.Generic.List[string]
    $bibliographyCodes = New-Object System.Collections.Generic.List[string]
    $bibliographyResults = New-Object System.Collections.Generic.List[string]
    $bibliographyResultLengths = New-Object System.Collections.Generic.List[int]
    $hyperlinkCount = 0

    for ($index = 1; $index -le $script:Document.Fields.Count; $index++) {
        $field = $script:Document.Fields.Item($index)
        try {
            $code = ''
            try { $code = [string]$field.Code.Text }
            catch {
                Write-Log -Level WARN -Message "Could not read the code of field ${index}: $($_.Exception.Message)"
                continue
            }
            if ($code -match $script:CitationDataFieldPattern) {
                $citationDataCodes.Add($code)
            }
            elseif ($code -match $script:CitationFieldPattern) {
                $citationCodes.Add($code)
                $citationResults.Add((Get-FieldResultText -Field $field -Label "citation field $index"))
            }
            elseif ($code -match $script:BibliographyFieldPattern) {
                $bibliographyCodes.Add($code)
                $bibliographyResult = Get-FieldResultText -Field $field -Label "bibliography field $index"
                $bibliographyResults.Add($bibliographyResult)
                $bibliographyResultLengths.Add($bibliographyResult.Trim().Length)
            }
            elseif ($code -match $script:HyperlinkFieldPattern) {
                $hyperlinkCount++
            }
        }
        finally { Release-ComObject $field }
    }

    return [ordered]@{
        citation_codes              = $citationCodes
        citation_results            = $citationResults
        citation_data_codes         = $citationDataCodes
        bibliography_codes          = $bibliographyCodes
        bibliography_results        = $bibliographyResults
        bibliography_result_lengths = $bibliographyResultLengths
        hyperlink_count             = $hyperlinkCount
    }
}

function Wait-CitationFieldCount {
    param([Parameter(Mandatory)][int]$Expected)
    $deadline = [DateTime]::UtcNow.AddMilliseconds($script:AutomationTimeoutMilliseconds)
    $fields = Get-CitationFields
    while ($fields.citation_codes.Count -lt $Expected -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
        $fields = Get-CitationFields
    }
    if ($fields.citation_codes.Count -lt $Expected) {
        throw "EndNote produced $($fields.citation_codes.Count) of $Expected expected ADDIN EN.CITE fields within $($script:AutomationTimeoutMilliseconds) ms. Raise ENDNOTE_AUTOMATION_TIMEOUT_MS, or check that the library is open and every rec_number exists."
    }
    return $fields
}

function Add-CoreValidationProblems {
    param($Fields, [string]$PlainText, $Manifest, $Errors)
    if ($PlainText.Contains($script:CitationMarkerPrefix)) {
        $Errors.Add("Unresolved citation marker text remains in the document ($($script:CitationMarkerPrefix)...).")
    }
    $residual = [regex]::Match($PlainText, $script:TemporaryCitationPattern)
    if ($residual.Success) {
        $Errors.Add("Unformatted temporary citation text remains in the document: $($residual.Value)")
    }
    $citationCount = $Fields.citation_codes.Count
    if ($null -ne $Manifest) {
        $expected = @(Get-OptionalProperty -Object $Manifest -Name 'citations').Count
        if ($citationCount -ne $expected) {
            $Errors.Add("Expected $expected native ADDIN EN.CITE fields; found $citationCount.")
        }
    }
    elseif ($citationCount -lt 1) { $Errors.Add('The document contains no native ADDIN EN.CITE field.') }
    foreach ($broken in $script:BrokenReferenceStrings) {
        if ($PlainText.Contains($broken)) {
            $Errors.Add("The document text contains a broken citation placeholder: $broken")
        }
    }
}

function Add-BibliographyValidationProblems {
    param($Fields, $Manifest, $Errors, $Warnings)
    if ($null -eq $Manifest) { return }
    $bibliography = Get-OptionalProperty -Object $Manifest -Name 'bibliography'
    $expectedFlag = Get-OptionalProperty -Object $bibliography -Name 'expected'
    if ($expectedFlag -isnot [bool]) {
        $Warnings.Add('The manifest does not declare bibliography.expected, so the bibliography gate was skipped.')
        return
    }
    if (-not [bool]$expectedFlag) { return }
    if ($Fields.bibliography_codes.Count -ne 1) {
        $Errors.Add("Expected exactly one ADDIN EN.REFLIST bibliography field; found $($Fields.bibliography_codes.Count).")
        return
    }
    if ($Fields.bibliography_result_lengths[0] -le 0) {
        $Errors.Add('The ADDIN EN.REFLIST bibliography field is empty.')
        return
    }
    $bibliographyText = [string]$Fields.bibliography_results[0]
    if ($bibliographyText -match $script:DuplicateDoiResolverPattern) {
        $Errors.Add('The bibliography contains a duplicated DOI resolver prefix.')
    }
    $checkedDois = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    foreach ($citation in @(Get-OptionalProperty -Object $Manifest -Name 'citations')) {
        foreach ($target in @(Get-OptionalProperty -Object $citation -Name 'targets')) {
            $doi = [string](Get-OptionalProperty -Object $target -Name 'doi')
            if (-not $checkedDois.Add($doi)) { continue }
            $visibleDoi = 'doi: ' + $script:DoiResolverPrefix + $doi
            if ($bibliographyText.IndexOf($visibleDoi, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                $Errors.Add("The bibliography does not visibly contain '$visibleDoi'. Check the output style Bibliography Templates.")
            }
        }
    }
}

function Add-TravelingLibraryValidationProblems {
    param($Payloads, $Manifest, $Errors, $Warnings)
    if ($null -eq $Manifest) {
        $Warnings.Add('No manifest supplied; record numbers, authors, style and library provenance were not checked.')
        return
    }
    foreach ($citation in @(Get-OptionalProperty -Object $Manifest -Name 'citations')) {
        $citationId = [string](Get-OptionalProperty -Object $citation -Name 'id')
        foreach ($target in @(Get-OptionalProperty -Object $citation -Name 'targets')) {
            $recNumber = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $target -Name 'rec_number') -Label "Citation '$citationId' rec_number"
            $author = ([string](Get-OptionalProperty -Object $target -Name 'author')).Trim()
            $element = [string]::Format($script:RecordNumberElementFormat, $recNumber)
            $recordSeen = $false
            $matched = $false
            foreach ($payload in $Payloads) {
                if ($payload.IndexOf($element, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                $recordSeen = $true
                if ($payload.IndexOf($author, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $matched = $true; break }
            }
            if (-not $recordSeen) {
                $Errors.Add("Citation '$citationId': rec-number $recNumber is absent from every traveling library payload.")
            }
            elseif (-not $matched) {
                $Errors.Add("Citation '$citationId': rec-number $recNumber was found, but the author surname '$author' does not appear in the same traveling library payload.")
            }
        }
    }
}

function Add-CitationDisplayValidationProblems {
    param($Fields, $Manifest, $Errors)
    if ($null -eq $Manifest) { return }
    foreach ($citation in @(Get-OptionalProperty -Object $Manifest -Name 'citations')) {
        if ((Get-CitationDisplay -Citation $citation) -ne $script:AuthorYearCitationDisplay) { continue }
        $citationId = [string](Get-OptionalProperty -Object $citation -Name 'id')
        $target = @(Get-OptionalProperty -Object $citation -Name 'targets')[0]
        $recNumber = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $target -Name 'rec_number') -Label "Citation '$citationId' rec_number"
        $author = [string](Get-OptionalProperty -Object $target -Name 'author')
        $element = [string]::Format($script:RecordNumberElementFormat, $recNumber)
        $displayMatched = $false
        for ($citationIndex = 0; $citationIndex -lt $Fields.citation_codes.Count; $citationIndex++) {
            $code = [string]$Fields.citation_codes[$citationIndex]
            $resultText = [string]$Fields.citation_results[$citationIndex]
            if ($code.IndexOf($element, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            if ($resultText.IndexOf($author, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $displayMatched = $true; break }
        }
        if (-not $displayMatched) {
            $Errors.Add("Citation '$citationId' requests display 'author-year', but no matching native field visibly contains author '$author'.")
        }
    }
}

function Get-ValidatedStyleName {
    param($Manifest, $Errors)
    $layout = Get-EndNoteLayout
    if ($null -eq $layout) {
        $Errors.Add("The $($script:LayoutVariableName) document variable is missing, so EndNote does not own this document's citations.")
        return $null
    }
    $styleName = $layout.style
    if ([string]::IsNullOrWhiteSpace($styleName)) {
        $Errors.Add("The $($script:LayoutVariableName) document variable carries no <Style> element.")
    }
    elseif ($null -ne $Manifest) {
        $manifestStyle = [string](Get-OptionalProperty -Object $Manifest -Name 'style')
        if ([string]::IsNullOrWhiteSpace($manifestStyle)) { $manifestStyle = Get-DefaultOutputStyle }
        if ($styleName -ine $manifestStyle) {
            $Errors.Add("The document output style is '$styleName' but the manifest requests '$manifestStyle'.")
        }
    }
    return $styleName
}

function Get-ValidatedLibraries {
    param($Manifest, $Errors, $Warnings)
    $libraries = @(Get-EndNoteLibraries -Warnings $Warnings)
    if ($libraries.Count -eq 0) {
        $Errors.Add("The $($script:LibrariesVariableName) document variable is missing or lists no library.")
        return $libraries
    }
    if ($null -eq $Manifest) { return $libraries }
    $recordedIds = New-Object System.Collections.Generic.HashSet[int]
    foreach ($library in $libraries) {
        foreach ($recordId in $library.record_ids) { [void]$recordedIds.Add([int]$recordId) }
    }
    foreach ($citation in @(Get-OptionalProperty -Object $Manifest -Name 'citations')) {
        $citationId = [string](Get-OptionalProperty -Object $citation -Name 'id')
        foreach ($target in @(Get-OptionalProperty -Object $citation -Name 'targets')) {
            $recNumber = ConvertTo-ManifestInteger -Value (Get-OptionalProperty -Object $target -Name 'rec_number') -Label "Citation '$citationId' rec_number"
            if (-not $recordedIds.Contains($recNumber)) {
                $Errors.Add("Citation '$citationId': rec-number $recNumber is absent from the $($script:LibrariesVariableName) record-ids list.")
            }
        }
    }
    return $libraries
}

function Get-ValidationReport {
    param(
        [Parameter(Mandatory)][string]$DocumentPath,
        $Manifest
    )
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $fields = Get-CitationFields
    $payloads = New-Object System.Collections.Generic.List[string]
    foreach ($code in $fields.citation_codes) { $payloads.Add($code) }
    foreach ($code in $fields.citation_data_codes) { $payloads.Add($code) }

    $plainText = [string]$script:Document.Content.Text
    $citationCount = $fields.citation_codes.Count
    Add-CoreValidationProblems -Fields $fields -PlainText $plainText -Manifest $Manifest -Errors $errors
    Add-BibliographyValidationProblems -Fields $fields -Manifest $Manifest -Errors $errors -Warnings $warnings
    Add-TravelingLibraryValidationProblems -Payloads $payloads -Manifest $Manifest -Errors $errors -Warnings $warnings
    Add-CitationDisplayValidationProblems -Fields $fields -Manifest $Manifest -Errors $errors
    $styleName = Get-ValidatedStyleName -Manifest $Manifest -Errors $errors
    $libraries = @(Get-ValidatedLibraries -Manifest $Manifest -Errors $errors -Warnings $warnings)

    return [ordered]@{
        ok            = $errors.Count -eq 0
        action        = 'validate'
        document_path = $DocumentPath
        counts        = [ordered]@{
            native_citation_fields = $citationCount
            citation_data_fields   = $fields.citation_data_codes.Count
            bibliography_fields    = $fields.bibliography_codes.Count
            hyperlink_fields       = $fields.hyperlink_count
        }
        style         = $styleName
        libraries     = @($libraries)
        errors        = @($errors)
        warnings      = @($warnings)
    }
}

# ---------------------------------------------------------------------------------------------
# Actions.
# ---------------------------------------------------------------------------------------------

function Invoke-Probe {
    $installation = $null
    try { $installation = Get-EndNoteInstallation }
    catch { Write-Log -Level WARN -Message "EndNote discovery failed: $($_.Exception.Message)" }

    $templatePath = $null
    try { $templatePath = Get-CwywTemplatePath }
    catch { Write-Log -Level WARN -Message "CWYW template discovery failed: $($_.Exception.Message)" }

    $addInRegistered = $false
    try { $addInRegistered = Test-Path -LiteralPath ($script:ProgIdRegistryPrefix + $script:CwywAddInProgId) }
    catch { Write-Log -Level WARN -Message "Could not read the CWYW ProgID registration: $($_.Exception.Message)" }

    $wordVersion = $null
    $templateLoaded = $false
    $macroCallable = $false
    try {
        # A throwaway empty document: the CWYW template only loads once a document is open, and
        # probing the macro here cannot touch user data.
        Start-Word
        $wordVersion = [string]$script:Word.Version
        foreach ($template in @($script:Word.Templates)) {
            if ($template.Name -like $script:CwywTemplateNamePattern) { $templateLoaded = $true }
            Release-ComObject $template
        }
        if ($templateLoaded) {
            try {
                Invoke-CwywFormat
                $macroCallable = $true
            }
            catch { Write-Log -Level WARN -Message "The CWYW format macro is not callable: $($_.Exception.Message)" }
        }
        else {
            Write-Log -Level WARN -Message 'The CWYW global template is not loaded, so the macro probe was skipped.'
        }
    }
    catch { Write-Log -Level WARN -Message "Word COM probe failed: $($_.Exception.Message)" }
    finally { Stop-Word }

    $styles = $null
    try { $styles = Get-InstalledStyles }
    catch { Write-Log -Level WARN -Message "Style discovery failed: $($_.Exception.Message)" }

    $endNoteRunning = $false
    try { $endNoteRunning = @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue).Count -gt 0 }
    catch { Write-Log -Level WARN -Message "Could not enumerate EndNote processes: $($_.Exception.Message)" }

    $executablePath = $null
    $productVersion = $null
    $stylesDirectory = $null
    if ($null -ne $installation) {
        $executablePath = $installation.executable
        $productVersion = $installation.product_version
        $stylesDirectory = $installation.styles_directory
    }
    $styleCount = 0
    if ($null -ne $styles) { $styleCount = @($styles.names).Count }

    $libraryPathValue = $null
    $libraryExists = $null
    $libraryDataExists = $null
    if (-not [string]::IsNullOrWhiteSpace($LibraryPath)) {
        # probe is the diagnostic surface and must still answer with exit 2 rather than crash
        # with exit 1 when the caller hands it a syntactically invalid path.
        try {
            $libraryPathValue = [IO.Path]::GetFullPath($LibraryPath)
            $libraryExists = Test-Path -LiteralPath $libraryPathValue -PathType Leaf
            $libraryDataExists = Test-Path -LiteralPath (Get-LibraryDataDirectory -LibraryFile $libraryPathValue) -PathType Container
        }
        catch {
            Write-Log -Level WARN -Message "Could not inspect the library path '$LibraryPath': $($_.Exception.Message)"
            $libraryPathValue = $LibraryPath
            $libraryExists = $false
            $libraryDataExists = $false
        }
    }

    $checks = [ordered]@{
        windows                  = $env:OS -eq 'Windows_NT'
        powershell               = $PSVersionTable.PSVersion.ToString()
        word_com                 = $null -ne $wordVersion
        word_version             = $wordVersion
        endnote_executable       = $null -ne $executablePath
        endnote_executable_path  = $executablePath
        endnote_product_version  = $productVersion
        cwyw_addin_registered    = $addInRegistered
        cwyw_addin_progid        = $script:CwywAddInProgId
        cwyw_template_present    = $null -ne $templatePath
        cwyw_template_path       = $templatePath
        cwyw_template_loaded     = $templateLoaded
        cwyw_macro_callable      = $macroCallable
        endnote_running          = $endNoteRunning
        styles_directory         = $stylesDirectory
        installed_style_count    = $styleCount
        library_path             = $libraryPathValue
        library_exists           = $libraryExists
        library_data_exists      = $libraryDataExists
    }

    $ready = $checks.windows -and
        $checks.word_com -and
        $checks.endnote_executable -and
        $checks.cwyw_addin_registered -and
        $checks.cwyw_template_present -and
        $checks.cwyw_template_loaded -and
        $checks.cwyw_macro_callable -and
        ($styleCount -gt 0)
    if ($null -ne $libraryPathValue) {
        $ready = $ready -and $libraryExists -and $libraryDataExists
    }

    Write-JsonResult ([ordered]@{
            ok     = $ready
            action = 'probe'
            checks = $checks
        })
    if (-not $ready) { exit 2 }
}

function Invoke-ListStyles {
    $styles = Get-InstalledStyles
    $names = @($styles.names)
    $total = $names.Count
    $filter = $null
    if (-not [string]::IsNullOrWhiteSpace($Style)) {
        $filter = $Style
        $escaped = [System.Management.Automation.WildcardPattern]::Escape($Style)
        $names = @($names | Where-Object { $_ -like "*$escaped*" })
    }
    $matched = $names.Count
    $truncated = $false
    if ($matched -gt $script:MaxStyleResults) {
        $names = @($names | Select-Object -First $script:MaxStyleResults)
        $truncated = $true
    }
    Write-JsonResult ([ordered]@{
            ok                    = $true
            action                = 'list-styles'
            styles_directories    = @($styles.directories)
            installed_style_count = $total
            filter                = $filter
            matched_style_count   = $matched
            truncated             = $truncated
            styles                = @($names)
        })
}

function Invoke-ConfigureDefaults {
    if (-not (Test-Path -LiteralPath $script:PackagedDefaultsPath -PathType Leaf)) {
        throw "Packaged defaults are missing: $script:PackagedDefaultsPath"
    }
    $defaults = $null
    try { $defaults = Get-Content -Raw -LiteralPath $script:PackagedDefaultsPath | ConvertFrom-Json }
    catch { throw "Packaged defaults are not valid JSON ($script:PackagedDefaultsPath): $($_.Exception.Message)" }

    New-Item -ItemType Directory -Path $script:UserConfigDirectory -Force | Out-Null
    Copy-Item -LiteralPath $script:PackagedDefaultsPath -Destination $script:UserDefaultsPath -Force
    Write-Log -Level INFO -Message "Wrote agent defaults to: $script:UserDefaultsPath"

    Write-JsonResult ([ordered]@{
            ok                            = $true
            action                        = 'configure-defaults'
            defaults_path                 = $script:UserDefaultsPath
            source_path                   = $script:PackagedDefaultsPath
            defaults                      = $defaults
            endnote_preferences_untouched = $true
            scope                         = "Defaults for every document processed by this bridge. EndNote's own registry preferences and UI state belong to the user and are never modified."
        })
}

function Get-ManifestTargetCount {
    param([Parameter(Mandatory)][object[]]$Citations)
    $count = 0
    foreach ($citation in $Citations) {
        $count += @(Get-OptionalProperty -Object $citation -Name 'targets').Count
    }
    return $count
}

# Resolve style, library and DOI gates before any GUI process starts. An unknown style,
# missing record or bad DOI is never allowed to reach EndNote's modal UI.
function Get-RenderContext {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'render requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'render requires OutputPath.' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath)) { throw 'render requires ManifestPath.' }

    $resolvedInput = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($resolvedInput) -ine $script:DocxExtension) { throw 'InputPath must be a .docx file.' }
    $output = Resolve-OutputFile -Path $OutputPath

    $resolvedManifest = Resolve-ExistingFile -Path $ManifestPath -Label 'ManifestPath'
    $manifest = Read-Manifest -Path $resolvedManifest
    Assert-Manifest -Manifest $manifest -LibraryOverride $LibraryPath

    $requestedStyle = [string](Get-OptionalProperty -Object $manifest -Name 'style')
    if ([string]::IsNullOrWhiteSpace($requestedStyle)) {
        $requestedStyle = Get-DefaultOutputStyle
        Write-Log -Level INFO -Message "Manifest omits style; using configured default '$requestedStyle'."
    }
    if (-not [string]::IsNullOrWhiteSpace($Style) -and $Style -ine $requestedStyle) {
        Write-Log -Level WARN -Message "Overriding the manifest style '$requestedStyle' with the -Style argument '$Style'."
        $requestedStyle = $Style
    }
    $styleName = Resolve-StyleName -Name $requestedStyle

    $libraryValue = $LibraryPath
    if ([string]::IsNullOrWhiteSpace($libraryValue)) {
        $libraryValue = [string](Get-OptionalProperty -Object (Get-OptionalProperty -Object $manifest -Name 'library') -Name 'path')
    }
    else {
        Write-Log -Level WARN -Message "Using the -LibraryPath argument '$LibraryPath' instead of the manifest library path."
    }
    $library = Assert-EndNoteLibrary -Path $libraryValue

    $precheck = Invoke-ManifestPrecheck -LibraryFile $library.path -ManifestFile $resolvedManifest
    $citations = @(Get-OptionalProperty -Object $manifest -Name 'citations')
    $targetCount = Get-ManifestTargetCount -Citations $citations
    return [ordered]@{
        input        = $resolvedInput
        output       = $output
        manifest     = $manifest
        manifest_path = $resolvedManifest
        style        = $styleName
        library      = $library
        precheck     = $precheck
        citations    = $citations
        target_count = $targetCount
    }
}

function Invoke-Render {
    $context = Get-RenderContext

    $endNoteProcessesBefore = Get-EndNoteProcessIds
    $endNote = Start-EndNoteWithLibrary -LibraryFile $context.library.path

    $fields = $null
    $authorYearCount = 0
    try {
        $workingCopy = New-WorkingCopy -SourcePath $context.input -Destination $context.output
        Start-Word -DocumentPath $workingCopy
        Write-Log -Level INFO -Message "Opened a working copy of: $($context.input)"
        Assert-CwywTemplateLoaded
        $null = Set-EndNoteLayout -StyleName $context.style

        $plans = Get-CitationMarkerPlans -Citations $context.citations
        Set-CitationMarkers -Plans $plans

        Invoke-CwywFormat
        $fields = Wait-CitationFieldCount -Expected $context.citations.Count
        $authorYearCount = Set-CitationDisplayFormats -Plans $plans -StyleName $context.style
        $fields = Wait-CitationFieldCount -Expected $context.citations.Count
        Save-DocumentAtomically -Destination $context.output
    }
    finally {
        Stop-Word
        Remove-WorkingCopy
    }
    Assert-EndNoteCleanupSucceeded

    Write-JsonResult ([ordered]@{
            ok                     = $true
            action                 = 'render'
            input_path             = $context.input
            output_path            = $context.output
            library_path           = $context.library.path
            style                  = $context.style
            citations              = $context.citations.Count
            author_year_citations  = $authorYearCount
            targets                = $context.target_count
            doi_verified           = @($context.precheck.resolved | Where-Object { $_.doi_verified }).Count
            native_citation_fields = $fields.citation_codes.Count
            citation_data_fields   = $fields.citation_data_codes.Count
            bibliography_present   = $fields.bibliography_codes.Count -gt 0
            endnote_preexisting    = $endNote.preexisting
            endnote_autostarted    = Test-EndNoteAutostarted -KnownProcessIds $endNoteProcessesBefore
            cleanup                = Get-EndNoteCleanupReport
        })
}

function Invoke-Validate {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'validate requires InputPath.' }
    $resolvedInput = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($resolvedInput) -ine $script:DocxExtension) { throw 'InputPath must be a .docx file.' }

    $manifest = $null
    if (-not [string]::IsNullOrWhiteSpace($ManifestPath)) {
        $manifest = Read-Manifest -Path $ManifestPath
        # The library file is irrelevant to the eight document gates and may legitimately live
        # on another machine, so its existence is not a validation precondition.
        Assert-Manifest -Manifest $manifest -SkipLibraryCheck
    }

    $report = $null
    try {
        # Fields are inspected exactly as stored. Fields.Update() is never called here because
        # updating an ADDIN EN.CITE field re-enters EndNote and can open a modal dialog.
        Start-Word -ReadOnly -DocumentPath $resolvedInput
        $report = Get-ValidationReport -DocumentPath $resolvedInput -Manifest $manifest
    }
    finally { Stop-Word }

    Write-JsonResult $report
    if (-not $report.ok) { exit 3 }
}

function Invoke-Update {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'update requires InputPath.' }
    $resolvedInput = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($resolvedInput) -ine $script:DocxExtension) { throw 'InputPath must be a .docx file.' }

    $destination = $resolvedInput
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $destination = Resolve-OutputFile -Path $OutputPath }
    if ($destination -ieq $resolvedInput -and -not $Overwrite) { throw 'Updating in place requires -Overwrite.' }

    $endNoteProcessesBefore = Get-EndNoteProcessIds
    $endNote = Confirm-EndNoteAvailable

    $fields = $null
    try {
        $workingCopy = New-WorkingCopy -SourcePath $resolvedInput -Destination $destination
        Start-Word -DocumentPath $workingCopy
        Invoke-CwywFormat
        $fields = Get-CitationFields
        Save-DocumentAtomically -Destination $destination
    }
    finally {
        Stop-Word
        Remove-WorkingCopy
    }
    Assert-EndNoteCleanupSucceeded

    Write-JsonResult ([ordered]@{
            ok                     = $true
            action                 = 'update'
            input_path             = $resolvedInput
            output_path            = $destination
            native_citation_fields = $fields.citation_codes.Count
            citation_data_fields   = $fields.citation_data_codes.Count
            bibliography_present   = $fields.bibliography_codes.Count -gt 0
            endnote_preexisting    = $endNote.preexisting
            endnote_autostarted    = Test-EndNoteAutostarted -KnownProcessIds $endNoteProcessesBefore
            cleanup                = Get-EndNoteCleanupReport
        })
}

function Invoke-SetStyle {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'set-style requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($Style)) { throw 'set-style requires Style.' }
    $resolvedInput = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($resolvedInput) -ine $script:DocxExtension) { throw 'InputPath must be a .docx file.' }

    $destination = $resolvedInput
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $destination = Resolve-OutputFile -Path $OutputPath }
    if ($destination -ieq $resolvedInput -and -not $Overwrite) { throw 'Changing the style in place requires -Overwrite.' }

    # Resolved before Word starts: an unknown style opens a modal EndNote dialog.
    $styleName = Resolve-StyleName -Name $Style
    $endNoteProcessesBefore = Get-EndNoteProcessIds
    $endNote = Confirm-EndNoteAvailable

    $previousStyle = $null
    $fields = $null
    try {
        $workingCopy = New-WorkingCopy -SourcePath $resolvedInput -Destination $destination
        Start-Word -DocumentPath $workingCopy
        Assert-CwywTemplateLoaded
        $layoutChange = Set-EndNoteLayout -StyleName $styleName
        $previousStyle = $layoutChange.previous_style
        Invoke-CwywFormat
        $fields = Get-CitationFields
        Save-DocumentAtomically -Destination $destination
    }
    finally {
        Stop-Word
        Remove-WorkingCopy
    }
    Assert-EndNoteCleanupSucceeded

    Write-JsonResult ([ordered]@{
            ok                     = $true
            action                 = 'set-style'
            input_path             = $resolvedInput
            output_path            = $destination
            previous_style         = $previousStyle
            style                  = $styleName
            native_citation_fields = $fields.citation_codes.Count
            citation_data_fields   = $fields.citation_data_codes.Count
            bibliography_present   = $fields.bibliography_codes.Count -gt 0
            endnote_preexisting    = $endNote.preexisting
            endnote_autostarted    = Test-EndNoteAutostarted -KnownProcessIds $endNoteProcessesBefore
            cleanup                = Get-EndNoteCleanupReport
        })
}

function Confirm-EndNoteAvailable {
    # Reformatting needs the library server. When a library path is supplied the bridge opens
    # EndNote with it; otherwise it only reports whether EndNote is already running, because
    # forcing a running EndNote to switch libraries would destroy the user's session.
    if (-not [string]::IsNullOrWhiteSpace($LibraryPath)) {
        $library = Assert-EndNoteLibrary -Path $LibraryPath
        return Start-EndNoteWithLibrary -LibraryFile $library.path
    }
    $running = @(Get-Process -Name $script:EndNoteProcessName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($running.Count -eq 0) {
        Write-Log -Level WARN -Message 'EndNote is not running and no -LibraryPath was supplied. Reformatting will only succeed for citations whose traveling library is embedded in the document.'
        return [ordered]@{ preexisting = $false; process_id = $null }
    }
    Write-Log -Level INFO -Message "EndNote is already running (PID $($running[0].Id))."
    return [ordered]@{ preexisting = $true; process_id = $running[0].Id }
}

function Invoke-Unformat {
    if ([string]::IsNullOrWhiteSpace($InputPath)) { throw 'unformat requires InputPath.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'unformat requires OutputPath; it never writes over its input.' }
    $resolvedInput = Resolve-ExistingFile -Path $InputPath -Label 'InputPath'
    if ([IO.Path]::GetExtension($resolvedInput) -ine $script:DocxExtension) { throw 'InputPath must be a .docx file.' }
    $output = Resolve-OutputFile -Path $OutputPath
    if ($output -ieq $resolvedInput) {
        throw 'unformat is destructive and always writes a new file. Supply an OutputPath different from InputPath; -Overwrite does not lift this rule.'
    }

    $fields = $null
    try {
        $workingCopy = New-WorkingCopy -SourcePath $resolvedInput -Destination $output
        Start-Word -DocumentPath $workingCopy
        Invoke-CwywMacro -Name $script:CwywUnformatMacro
        $fields = Get-CitationFields
        Save-DocumentAtomically -Destination $output
    }
    finally {
        Stop-Word
        Remove-WorkingCopy
    }
    Assert-EndNoteCleanupSucceeded

    Write-JsonResult ([ordered]@{
            ok                                = $true
            action                            = 'unformat'
            input_path                        = $resolvedInput
            output_path                       = $output
            remaining_native_citation_fields  = $fields.citation_codes.Count
            remaining_bibliography_fields     = $fields.bibliography_codes.Count
            note                              = 'Formatted citations were reverted to temporary {Author, Year #RecNum} text. Run the update action to format them again.'
            cleanup                           = Get-EndNoteCleanupReport
        })
}

function Invoke-ImportRis {
    if ([string]::IsNullOrWhiteSpace($LibraryPath)) { throw 'import-ris requires LibraryPath.' }
    if ([string]::IsNullOrWhiteSpace($RisPath)) { throw 'import-ris requires RisPath.' }
    $resolvedRis = Resolve-ExistingFile -Path $RisPath -Label 'RisPath'
    if ([IO.Path]::GetExtension($resolvedRis) -ine $script:RisExtension) {
        throw "RisPath must use the $($script:RisExtension) extension: $resolvedRis"
    }
    $strictUtf8 = [System.Text.UTF8Encoding]::new($false, $true)
    $risText = [IO.File]::ReadAllText($resolvedRis, $strictUtf8)
    $risRecordCount = [regex]::Matches($risText, '(?m)^TY  - ').Count
    if ($risRecordCount -lt 1) {
        throw 'RisPath contains no RIS records (no line beginning with TY  - ).'
    }

    $resolved = Resolve-EndNoteImportLibrary -TargetPath $LibraryPath -EmptyTemplatePath $EmptyLibraryTemplatePath
    $library = $resolved.library
    $beforeCount = Get-EndNoteRecordCount -LibraryFile $library.path
    $installation = Get-EndNoteInstallation
    if ($null -eq $installation) {
        throw 'EndNote is not installed or could not be discovered, so RIS cannot be imported through EndNote Direct Export.'
    }
    $opened = Open-EndNoteLibraryForImport -LibraryFile $library.path -Installation $installation

    Write-Log -Level INFO -Message "Sending RIS to EndNote Direct Export: $resolvedRis"
    Start-Process -FilePath $installation.executable -ArgumentList ('"' + $resolvedRis + '"') -WindowStyle Hidden | Out-Null

    $afterCount = $beforeCount
    $deadline = [DateTime]::UtcNow.AddMilliseconds($script:EndNoteReadyTimeoutMilliseconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds $script:AutomationPollMilliseconds
        try { $afterCount = Get-EndNoteRecordCount -LibraryFile $library.path }
        catch {
            # EndNote briefly locks or rotates the SQLite files while committing. Retry until
            # the deadline rather than treating that expected transient as import failure.
            continue
        }
        if ($afterCount -gt $beforeCount) { break }
    }

    $importedCount = $afterCount - $beforeCount
    $note = $null
    if ($importedCount -eq 0) {
        $note = 'The live-reference count did not increase after EndNote Direct Export. Search the exact library by DOI: all records may already exist, or EndNote may have rejected or discarded the import.'
        Write-Log -Level WARN -Message $note
    }
    Write-JsonResult ([ordered]@{
            ok                          = ($importedCount -gt 0)
            action                      = 'import-ris'
            library_path                = $library.path
            library_created             = [bool]$resolved.created
            empty_library_template_path = $resolved.template_path
            ris_path                    = $resolvedRis
            ris_record_count             = $risRecordCount
            live_references_before       = $beforeCount
            live_references_after        = $afterCount
            imported_count               = $importedCount
            endnote_preexisting_target   = [bool]$opened.preexisting
            endnote_window_title         = $opened.title
            confirmed_library_path       = $opened.confirmed_library_path
            import_route                 = 'EndNote Direct Export'
            note                         = $note
        })
    if ($importedCount -eq 0) { exit 3 }
}

# ---------------------------------------------------------------------------------------------
# Entry point.
# ---------------------------------------------------------------------------------------------

try {
    switch ($Action) {
        'probe' { Invoke-Probe }
        'list-styles' { Invoke-ListStyles }
        'configure-defaults' { Invoke-ConfigureDefaults }
        'render' { Invoke-Render }
        'validate' { Invoke-Validate }
        'update' { Invoke-Update }
        'set-style' { Invoke-SetStyle }
        'unformat' { Invoke-Unformat }
        'import-ris' { Invoke-ImportRis }
    }
}
catch {
    # The .NET stack alone points at the CLR frame, not at the bridge line that failed, which
    # makes a type-conversion error inside a COM loop practically undiagnosable. Record the
    # PowerShell frame as well and hand the caller the originating line.
    $scriptStack = ''
    $errorLocation = $null
    try {
        $scriptStack = [string]$_.ScriptStackTrace
        $errorLocation = [string]$_.InvocationInfo.PositionMessage
    }
    catch {
        Write-Log -Level WARN -Message "Could not read the PowerShell stack for the failing command: $($_.Exception.Message)"
    }
    Write-Log -Level ERROR -Message $_.Exception.ToString()
    if (-not [string]::IsNullOrWhiteSpace($scriptStack)) {
        Write-Log -Level ERROR -Message "PowerShell stack:`n$scriptStack"
    }
    Write-JsonResult ([ordered]@{
            ok             = $false
            action         = $Action
            error          = $_.Exception.Message
            error_location = $errorLocation
            script_stack   = $scriptStack
            cleanup_errors = @($script:CleanupFailures)
        })
    exit 1
}
finally {
    Stop-Word
    Remove-PendingTemporarySave
    Remove-WorkingCopy
}
