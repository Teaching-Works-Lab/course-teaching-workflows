[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [switch]$Json
)

$ErrorActionPreference = "Stop"
$inputFullPath = [System.IO.Path]::GetFullPath($InputPath)
$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)

if (-not (Test-Path -LiteralPath $inputFullPath -PathType Leaf)) {
    throw "Legacy DOC input does not exist: $inputFullPath"
}

$outputDirectory = [System.IO.Path]::GetDirectoryName($outputFullPath)
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    [System.IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
}

$word = $null
$document = $null
$wordVersion = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputFullPath, $false, $true, $false)
    try {
        $document.SaveAs2($outputFullPath, 16)
    }
    catch {
        $docxFormat = 16
        $document.SaveAs([ref]$outputFullPath, [ref]$docxFormat)
    }

    if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
        throw "Microsoft Word did not create the requested DOCX output: $outputFullPath"
    }
    $wordVersion = [string]$word.Version
}
finally {
    if ($null -ne $document) {
        try {
            $document.Close(0)
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
    }
    if ($null -ne $word) {
        try {
            $word.Quit()
        }
        catch {
            if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
                throw
            }
        }
        finally {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
    }
}

if (-not (Test-Path -LiteralPath $outputFullPath -PathType Leaf)) {
    throw "Microsoft Word did not create the requested DOCX output: $outputFullPath"
}

$result = [ordered]@{
    tool = "Microsoft Word"
    version = $wordVersion
}
if ($Json) {
    $result | ConvertTo-Json -Compress
}
else {
    Write-Output "$($result.tool) $($result.version)"
}
