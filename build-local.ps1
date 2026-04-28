param(
  [switch]$Clean
)

$ErrorActionPreference = "Stop"

$previousPublishMode = $env:PUBLISH_MODE
$rootDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Push-Location $rootDir
try {
  if (-not (Test-Path "node_modules")) {
    Write-Host "Installiere Abhaengigkeiten mit npm ci..."
    npm ci
  }

  if ($Clean -and (Test-Path "builds")) {
    Write-Host "Entferne alten builds-Ordner..."
    Remove-Item -LiteralPath "builds" -Recurse -Force
  }

  Write-Host "Baue Tailwind CSS..."
  npm run build:css

  $env:PUBLISH_MODE = "never"
  Write-Host "Erzeuge lokalen Windows-Build..."
  node ".\scripts\release-with-icon.js"

  Write-Host ""
  Write-Host "Fertig. Installer liegt im builds-Ordner:"
  Get-ChildItem ".\builds\Repetierer-*.exe" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 5 Name, Length, LastWriteTime |
    Format-Table -AutoSize
}
finally {
  if ($null -eq $previousPublishMode) {
    Remove-Item Env:PUBLISH_MODE -ErrorAction SilentlyContinue
  }
  else {
    $env:PUBLISH_MODE = $previousPublishMode
  }

  Pop-Location
}
