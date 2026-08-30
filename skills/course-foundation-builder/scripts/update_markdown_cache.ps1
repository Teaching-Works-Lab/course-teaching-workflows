param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $true)]
    [string]$CacheRoot,

    [string]$MarkItDownPath = "markitdown",
    [string]$PandocPath = "pandoc",
    [switch]$Force,
    [switch]$Json
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$sourceRootPath = [System.IO.Path]::GetFullPath($SourceRoot).TrimEnd('\', '/')
$cacheRootPath = [System.IO.Path]::GetFullPath($CacheRoot).TrimEnd('\', '/')
$markdownRootPath = Join-Path $cacheRootPath "by-source"
$manifestPath = Join-Path $cacheRootPath "manifest.json"
$logPath = Join-Path $cacheRootPath "conversion-log.jsonl"

if (-not (Test-Path -LiteralPath $sourceRootPath -PathType Container)) {
    Write-Error "Source directory does not exist: $sourceRootPath"
    exit 1
}

[System.IO.Directory]::CreateDirectory($markdownRootPath) | Out-Null

function Get-RelativeSourcePath {
    param([string]$FullPath)
    return $FullPath.Substring($sourceRootPath.Length).TrimStart('\', '/')
}

function Get-ToolVersion {
    param([string]$Command)
    try {
        $line = & $Command --version 2>$null | Select-Object -First 1
        if ([string]::IsNullOrWhiteSpace($line)) { return "unknown" }
        return [string]$line
    }
    catch {
        return "unavailable"
    }
}

function Get-Sha256 {
    param([string]$Path)
    $stream = [System.IO.File]::OpenRead($Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
        $stream.Dispose()
    }
}

function Convert-LegacyDocToDocx {
    param([string]$InputPath, [string]$OutputPath)
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Open($InputPath, $false, $true, $false)
        try {
            $document.SaveAs2($OutputPath, 16)
        }
        catch {
            $document.SaveAs([ref]$OutputPath, [ref]16)
        }
        return [string]$word.Version
    }
    finally {
        if ($null -ne $document) {
            $document.Close(0)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch { if (-not (Test-Path -LiteralPath $OutputPath)) { throw } }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
    }
}

$existingBySource = @{}
if (Test-Path -LiteralPath $manifestPath) {
    $existingManifest = [System.IO.File]::ReadAllText($manifestPath, $utf8NoBom) | ConvertFrom-Json
    foreach ($entry in @($existingManifest.files)) {
        $existingBySource[[string]$entry.source_relative_path] = $entry
    }
}

$supported = @(".md", ".doc", ".docx", ".pdf", ".html", ".htm", ".epub", ".pptx", ".xlsx")
$cachePrefix = $cacheRootPath + [System.IO.Path]::DirectorySeparatorChar
$sourceFiles = Get-ChildItem -LiteralPath $sourceRootPath -Recurse -File |
    Where-Object {
        $full = [System.IO.Path]::GetFullPath($_.FullName)
        -not $full.StartsWith($cachePrefix, [System.StringComparison]::OrdinalIgnoreCase) -and
        $_.Extension.ToLowerInvariant() -in $supported
    } |
    Sort-Object FullName

$tempRootPath = Join-Path ([System.IO.Path]::GetTempPath()) ("course-foundation-cache-" + [guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory($tempRootPath) | Out-Null
$results = New-Object System.Collections.Generic.List[object]
$errors = New-Object System.Collections.Generic.List[string]
$convertedCount = 0
$reusedCount = 0
$reviewCount = 0

try {
    foreach ($sourceFile in $sourceFiles) {
        $relativePath = Get-RelativeSourcePath $sourceFile.FullName
        $relativeDirectory = [System.IO.Path]::GetDirectoryName($relativePath)
        $outputDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
            $markdownRootPath
        } else {
            Join-Path $markdownRootPath $relativeDirectory
        }
        [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null

        $extension = $sourceFile.Extension.ToLowerInvariant()
        $outputFileName = if ($extension -eq ".md") { $sourceFile.Name } else { $sourceFile.Name + ".md" }
        $outputPath = Join-Path $outputDirectory $outputFileName
        $cachePath = Get-RelativeSourcePath $outputPath
        if ($outputPath.StartsWith($cacheRootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $cachePath = $outputPath.Substring($cacheRootPath.Length).TrimStart('\', '/')
        }
        $sha256 = Get-Sha256 $sourceFile.FullName
        $existing = $existingBySource[$relativePath]

        if (-not $Force -and $null -ne $existing -and $existing.sha256 -eq $sha256 -and (Test-Path -LiteralPath $outputPath)) {
            $reusedCount += 1
            $results.Add([pscustomobject]@{
                source_relative_path = $relativePath
                extension = $extension
                sha256 = $sha256
                source_last_write_time_utc = $sourceFile.LastWriteTimeUtc.ToString("o")
                route = $existing.route
                tool = $existing.tool
                tool_version = $existing.tool_version
                status = "reused"
                cache_path = $cachePath
                converted_at_utc = $existing.converted_at_utc
            })
            continue
        }

        try {
            $route = ""
            $tool = ""
            $toolVersion = ""
            $status = "converted"

            switch ($extension) {
                ".md" {
                    $route = "direct-markdown"
                    $tool = "existing-markdown"
                    $toolVersion = "source"
                    [System.IO.File]::Copy($sourceFile.FullName, $outputPath, $true)
                }
                ".html" {
                    $route = "pandoc-html"
                    $tool = "pandoc"
                    $toolVersion = Get-ToolVersion $PandocPath
                    & $PandocPath $sourceFile.FullName -t gfm -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "Pandoc conversion failed" }
                }
                ".htm" {
                    $route = "pandoc-html"
                    $tool = "pandoc"
                    $toolVersion = Get-ToolVersion $PandocPath
                    & $PandocPath $sourceFile.FullName -t gfm -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "Pandoc conversion failed" }
                }
                ".epub" {
                    $route = "pandoc-epub"
                    $tool = "pandoc"
                    $toolVersion = Get-ToolVersion $PandocPath
                    & $PandocPath $sourceFile.FullName -t gfm -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "Pandoc conversion failed" }
                }
                ".doc" {
                    $route = "word-doc-to-docx+markitdown"
                    $tool = "word+markitdown"
                    $temporaryDocxPath = Join-Path $tempRootPath ([guid]::NewGuid().ToString("N") + ".docx")
                    $wordVersion = Convert-LegacyDocToDocx $sourceFile.FullName $temporaryDocxPath
                    $toolVersion = "Word $wordVersion; " + (Get-ToolVersion $MarkItDownPath)
                    & $MarkItDownPath $temporaryDocxPath -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "MarkItDown conversion failed" }
                }
                ".docx" {
                    $route = "markitdown-docx"
                    $tool = "markitdown"
                    $toolVersion = Get-ToolVersion $MarkItDownPath
                    & $MarkItDownPath $sourceFile.FullName -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "MarkItDown conversion failed" }
                }
                ".pdf" {
                    $route = "markitdown-pdf"
                    $tool = "markitdown"
                    $toolVersion = Get-ToolVersion $MarkItDownPath
                    & $MarkItDownPath $sourceFile.FullName -o $outputPath | Out-Null
                    if ($LASTEXITCODE -ne 0) { throw "MarkItDown conversion failed" }
                    $status = "converted-review-recommended"
                    $reviewCount += 1
                }
                default {
                    $route = "artifact-tool-required"
                    $tool = "document-artifact-tool"
                    $toolVersion = "not-run"
                    $status = "needs-extraction"
                    $reviewCount += 1
                    [System.IO.File]::WriteAllText(
                        $outputPath,
                        "---`nsource_file: '$relativePath'`nstatus: needs-extraction`n---`n",
                        $utf8NoBom
                    )
                }
            }

            if (-not (Test-Path -LiteralPath $outputPath)) { throw "Expected cache output was not created" }
            $convertedCount += 1
            $results.Add([pscustomobject]@{
                source_relative_path = $relativePath
                extension = $extension
                sha256 = $sha256
                source_last_write_time_utc = $sourceFile.LastWriteTimeUtc.ToString("o")
                route = $route
                tool = $tool
                tool_version = $toolVersion
                status = $status
                cache_path = $cachePath
                converted_at_utc = [DateTime]::UtcNow.ToString("o")
            })
        }
        catch {
            $message = "${relativePath}: $($_.Exception.Message)"
            $errors.Add($message)
            $results.Add([pscustomobject]@{
                source_relative_path = $relativePath
                extension = $extension
                sha256 = $sha256
                source_last_write_time_utc = $sourceFile.LastWriteTimeUtc.ToString("o")
                route = "failed"
                tool = "unknown"
                tool_version = "unknown"
                status = "failed"
                cache_path = $null
                converted_at_utc = [DateTime]::UtcNow.ToString("o")
                error = $_.Exception.Message
            })
        }
    }
}
finally {
    if (Test-Path -LiteralPath $tempRootPath) {
        $resolvedTemp = [System.IO.Path]::GetFullPath($tempRootPath)
        $systemTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if ($resolvedTemp.StartsWith($systemTemp, [System.StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
        }
    }
}

$manifest = [ordered]@{
    schema_version = 1
    source_root = $sourceRootPath
    cache_root = $cacheRootPath
    generated_at_utc = [DateTime]::UtcNow.ToString("o")
    source_count = $sourceFiles.Count
    files = $results
}
[System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 8), $utf8NoBom)

$summary = [ordered]@{
    run_id = [guid]::NewGuid().ToString("N")
    generated_at_utc = [DateTime]::UtcNow.ToString("o")
    source_root = $sourceRootPath
    cache_root = $cacheRootPath
    converted = $convertedCount
    reused = $reusedCount
    requires_review = $reviewCount
    failed = $errors.Count
    errors = @($errors)
}
$summaryJson = $summary | ConvertTo-Json -Compress -Depth 6
[System.IO.File]::AppendAllText($logPath, $summaryJson + [Environment]::NewLine, $utf8NoBom)

if ($Json) {
    Write-Output $summaryJson
} else {
    Write-Output "Cache root: $cacheRootPath"
    Write-Output "Converted: $convertedCount"
    Write-Output "Reused: $reusedCount"
    Write-Output "Requires review: $reviewCount"
    Write-Output "Failed: $($errors.Count)"
}

if ($errors.Count -gt 0) { exit 1 }
if ($reviewCount -gt 0) { exit 2 }
exit 0
