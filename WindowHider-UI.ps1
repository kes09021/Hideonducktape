[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$ValidateOnly,
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$partsDirectory = Join-Path -Path $PSScriptRoot -ChildPath 'src'
if (-not (Test-Path -LiteralPath $partsDirectory)) {
    throw "UI source parts were not found in '$partsDirectory'."
}

$parts = @(
    Get-ChildItem -LiteralPath $partsDirectory -Filter 'WindowHider-UI.part*.ps1' |
        Sort-Object Name
)

if ($parts.Count -eq 0) {
    throw "No UI source parts were found in '$partsDirectory'."
}

$generatedScriptName = '.WindowHider-UI.generated.{0}.ps1' -f ([System.Guid]::NewGuid().ToString('N'))
$generatedScriptPath = Join-Path -Path $PSScriptRoot -ChildPath $generatedScriptName
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$scriptText = (
    $parts |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
) -join [Environment]::NewLine

[System.IO.File]::WriteAllText($generatedScriptPath, $scriptText, $utf8NoBom)

try {
    $forwardParameters = @{}
    if ($PSBoundParameters.ContainsKey('ConfigPath')) {
        $forwardParameters.ConfigPath = $ConfigPath
    }

    if ($ValidateOnly) {
        $forwardParameters.ValidateOnly = $true
    }

    if ($SmokeTest) {
        $forwardParameters.SmokeTest = $true
    }

    & $generatedScriptPath @forwardParameters
    if ($LASTEXITCODE -is [int]) {
        exit $LASTEXITCODE
    }
}
finally {
    Remove-Item -LiteralPath $generatedScriptPath -ErrorAction SilentlyContinue
}
