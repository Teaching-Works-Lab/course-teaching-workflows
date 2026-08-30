[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $true)]
    [string]$CacheRoot,

    [string]$MarkItDownPath,
    [string]$PandocPath,
    [string]$LibreOfficePath,
    [switch]$Force,
    [switch]$Json
)

$ErrorActionPreference = "Stop"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$pythonCandidates = @(
    @{ Command = 'py'; Prefix = @('-3') },
    @{ Command = 'python3'; Prefix = @() },
    @{ Command = 'python'; Prefix = @() }
)

$pythonCommand = $null
$pythonPrefix = @()
foreach ($candidate in $pythonCandidates) {
    $resolved = Get-Command $candidate.Command -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $resolved) {
        $pythonCommand = $resolved.Source
        $pythonPrefix = @($candidate.Prefix)
        break
    }
}

if ($null -eq $pythonCommand) {
    Write-Error "Python 3 is required."
    exit 1
}

$scriptPath = Join-Path $PSScriptRoot "update_markdown_cache.py"
$pythonArguments = @(
    $pythonPrefix
    $scriptPath
    "--source-root"
    $SourceRoot
    "--cache-root"
    $CacheRoot
)

if ($PSBoundParameters.ContainsKey("MarkItDownPath")) {
    $pythonArguments += @("--markitdown-command", $MarkItDownPath)
}
if ($PSBoundParameters.ContainsKey("PandocPath")) {
    $pythonArguments += @("--pandoc-command", $PandocPath)
}
if ($PSBoundParameters.ContainsKey("LibreOfficePath")) {
    $pythonArguments += @("--libreoffice-command", $LibreOfficePath)
}
if ($Force) {
    $pythonArguments += "--force"
}
if ($Json) {
    $pythonArguments += "--json"
}

$nativeErrorPreference = Get-Variable -Name PSNativeCommandUseErrorActionPreference `
    -ErrorAction SilentlyContinue
if ($null -ne $nativeErrorPreference) {
    $previousNativeErrorPreference = $nativeErrorPreference.Value
    Set-Variable -Name PSNativeCommandUseErrorActionPreference -Value $false
}
try {
    & $pythonCommand @pythonArguments
    $pythonExitCode = $LASTEXITCODE
}
finally {
    if ($null -ne $nativeErrorPreference) {
        Set-Variable -Name PSNativeCommandUseErrorActionPreference `
            -Value $previousNativeErrorPreference
    }
}
exit $pythonExitCode
