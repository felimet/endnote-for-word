function Get-EndNoteNormalizedPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
}

function Test-EndNoteSamePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Left,

        [Parameter(Mandatory = $true)]
        [string]$Right
    )

    $leftPath = Get-EndNoteNormalizedPath -Path $Left
    $rightPath = Get-EndNoteNormalizedPath -Path $Right
    return [string]::Equals($leftPath, $rightPath, [System.StringComparison]::OrdinalIgnoreCase)
}

function Assert-EndNoteGeneratedDocumentPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    $fullPath = Get-EndNoteNormalizedPath -Path $Path
    $fullDestination = Get-EndNoteNormalizedPath -Path $Destination
    $artifactDirectory = Split-Path -Parent $fullPath
    $destinationDirectory = Split-Path -Parent $fullDestination

    if (-not (Test-EndNoteSamePath -Left $artifactDirectory -Right $destinationDirectory)) {
        throw "Refusing to remove generated document outside its destination directory: $fullPath"
    }

    $baseName = [regex]::Escape([System.IO.Path]::GetFileNameWithoutExtension($fullDestination))
    $expectedName = "(?i)^\.$baseName\.[0-9a-f]{32}\.(work|tmp)\.docx$"
    if ([System.IO.Path]::GetFileName($fullPath) -notmatch $expectedName) {
        throw "Refusing to remove a file that does not match the generated document pattern: $fullPath"
    }

    return $fullPath
}

function Remove-EndNoteGeneratedDocumentFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    $verifiedPath = Assert-EndNoteGeneratedDocumentPath -Path $Path -Destination $Destination
    if (-not (Test-Path -LiteralPath $verifiedPath)) {
        return $false
    }

    $item = Get-Item -LiteralPath $verifiedPath -Force -ErrorAction Stop
    if ($item.PSIsContainer -or (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw "Refusing to remove a directory or reparse point as a generated document: $verifiedPath"
    }

    Remove-Item -LiteralPath $verifiedPath -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $verifiedPath) {
        throw "Cleanup verification failed; generated document still exists: $verifiedPath"
    }

    return $true
}

function Remove-EndNoteStaleDocumentFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Destination,

        [int]$OlderThanHours = 24
    )

    if ($OlderThanHours -lt 1) {
        throw 'OlderThanHours must be at least 1.'
    }

    $fullDestination = Get-EndNoteNormalizedPath -Path $Destination
    $directory = Split-Path -Parent $fullDestination
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        return @()
    }

    $baseName = [regex]::Escape([System.IO.Path]::GetFileNameWithoutExtension($fullDestination))
    $expectedName = "(?i)^\.$baseName\.[0-9a-f]{32}\.(work|tmp)\.docx$"
    # ponytail: 24 hours protects concurrent edits while bounding abandoned bridge artifacts.
    $cutoff = [DateTime]::UtcNow.AddHours(-$OlderThanHours)
    $removed = New-Object 'System.Collections.Generic.List[string]'

    foreach ($item in @(Get-ChildItem -LiteralPath $directory -File -Force -ErrorAction Stop)) {
        if ($item.Name -notmatch $expectedName -or $item.LastWriteTimeUtc -gt $cutoff) {
            continue
        }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            continue
        }

        if (Remove-EndNoteGeneratedDocumentFile -Path $item.FullName -Destination $fullDestination) {
            $removed.Add($item.FullName) | Out-Null
        }
    }

    return @($removed)
}

function Assert-EndNoteGeneratedDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Parent,

        [Parameter(Mandatory = $true)]
        [ValidateSet('endnote-tests', 'endnote-live')]
        [string]$Prefix
    )

    $fullPath = Get-EndNoteNormalizedPath -Path $Path
    $fullParent = Get-EndNoteNormalizedPath -Path $Parent
    if (-not (Test-EndNoteSamePath -Left (Split-Path -Parent $fullPath) -Right $fullParent)) {
        throw "Refusing to remove generated directory outside its expected parent: $fullPath"
    }

    $expectedName = "(?i)^$([regex]::Escape($Prefix))-[0-9a-f]{32}$"
    if ([System.IO.Path]::GetFileName($fullPath) -notmatch $expectedName) {
        throw "Refusing to remove a directory that does not match the generated test pattern: $fullPath"
    }

    return $fullPath
}

function Assert-EndNoteDirectoryTreeHasNoReparsePoints {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    foreach ($item in @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop)) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing recursive cleanup because the generated directory contains a reparse point: $($item.FullName)"
        }
    }
}

function Remove-EndNoteGeneratedDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Parent,

        [Parameter(Mandatory = $true)]
        [ValidateSet('endnote-tests', 'endnote-live')]
        [string]$Prefix
    )

    $verifiedPath = Assert-EndNoteGeneratedDirectory -Path $Path -Parent $Parent -Prefix $Prefix
    if (-not (Test-Path -LiteralPath $verifiedPath)) {
        return $false
    }

    $item = Get-Item -LiteralPath $verifiedPath -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw "Refusing to recursively remove a non-directory or reparse point: $verifiedPath"
    }

    Assert-EndNoteDirectoryTreeHasNoReparsePoints -Path $verifiedPath
    Remove-Item -LiteralPath $verifiedPath -Recurse -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $verifiedPath) {
        throw "Cleanup verification failed; generated directory still exists: $verifiedPath"
    }

    return $true
}

function Remove-EndNoteStaleTestDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Parent,

        [Parameter(Mandatory = $true)]
        [ValidateSet('endnote-tests', 'endnote-live')]
        [string]$Prefix,

        [int]$OlderThanHours = 24
    )

    if ($OlderThanHours -lt 1) {
        throw 'OlderThanHours must be at least 1.'
    }

    $fullParent = Get-EndNoteNormalizedPath -Path $Parent
    if (-not (Test-Path -LiteralPath $fullParent -PathType Container)) {
        return @()
    }

    # ponytail: 24 hours avoids deleting a concurrent test run while sweeping abandoned runs.
    $cutoff = [DateTime]::UtcNow.AddHours(-$OlderThanHours)
    $expectedName = "(?i)^$([regex]::Escape($Prefix))-[0-9a-f]{32}$"
    $removed = New-Object 'System.Collections.Generic.List[string]'

    foreach ($item in @(Get-ChildItem -LiteralPath $fullParent -Directory -Force -ErrorAction Stop)) {
        if ($item.Name -notmatch $expectedName -or $item.LastWriteTimeUtc -gt $cutoff) {
            continue
        }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            continue
        }

        $treeItems = @(Get-ChildItem -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop)
        if (@($treeItems | Where-Object {
                    ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0 -or
                    $_.LastWriteTimeUtc -gt $cutoff
                }).Count -gt 0) {
            continue
        }

        if (Remove-EndNoteGeneratedDirectory -Path $item.FullName -Parent $fullParent -Prefix $Prefix) {
            $removed.Add($item.FullName) | Out-Null
        }
    }

    return @($removed)
}

function Remove-EndNoteRepoCacheDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('scripts\__pycache__', 'tests\__pycache__')]
        [string]$RelativePath
    )

    $fullRoot = Get-EndNoteNormalizedPath -Path $RepositoryRoot
    $cachePath = Get-EndNoteNormalizedPath -Path (Join-Path $fullRoot $RelativePath)
    $expectedPath = Get-EndNoteNormalizedPath -Path (Join-Path $fullRoot $RelativePath)
    if (-not (Test-EndNoteSamePath -Left $cachePath -Right $expectedPath)) {
        throw "Refusing to remove an unexpected repository cache path: $cachePath"
    }
    if (-not (Test-Path -LiteralPath $cachePath)) {
        return $false
    }

    $item = Get-Item -LiteralPath $cachePath -Force -ErrorAction Stop
    if (-not $item.PSIsContainer -or (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0)) {
        throw "Refusing to remove a non-directory or reparse point as a repository cache: $cachePath"
    }

    Assert-EndNoteDirectoryTreeHasNoReparsePoints -Path $cachePath
    Remove-Item -LiteralPath $cachePath -Recurse -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $cachePath) {
        throw "Cleanup verification failed; repository cache still exists: $cachePath"
    }

    return $true
}
