param(
  [string]$Notes = "更新系統版本"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$encoding = New-Object System.Text.UTF8Encoding($false)
$today = Get-Date -Format "yyyy.MM.dd"
$versionPath = Join-Path $root "version.json"
$sequence = 1

if (Test-Path $versionPath) {
  $oldVersion = (Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json).version
  if ($oldVersion -match "^$([regex]::Escape($today))-(\d+)$") {
    $sequence = [int]$Matches[1] + 1
  }
}

$version = "$today-$sequence"
$versionJson = [ordered]@{
  version = $version
  notes = $Notes
} | ConvertTo-Json
[System.IO.File]::WriteAllText($versionPath, $versionJson + [Environment]::NewLine, $encoding)

$swPath = Join-Path $root "sw.js"
$swText = [System.IO.File]::ReadAllText($swPath, [System.Text.Encoding]::UTF8)
$swText = [regex]::Replace($swText, "const BUILD_VERSION = '[^']+';", "const BUILD_VERSION = '$version';")
[System.IO.File]::WriteAllText($swPath, $swText, $encoding)

Write-Host "版本已更新：$version"
