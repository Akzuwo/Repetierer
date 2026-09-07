param(
  [switch]$Clean
)

$ErrorActionPreference = "Stop"

$previousPublishMode = $env:PUBLISH_MODE
$previousLocalBuildNumber = $env:LOCAL_BUILD_NUMBER
$rootDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Invoke-NativeCommand {
  param(
    [Parameter(Mandatory = $true)][scriptblock]$Command,
    [Parameter(Mandatory = $true)][string]$Description
  )

  & $Command
  if ($LASTEXITCODE -ne 0) {
    throw "$Description fehlgeschlagen (Exitcode $LASTEXITCODE)."
  }
}

Push-Location $rootDir
try {
  if (-not (Test-Path "node_modules")) {
    Write-Host "Installiere Abhaengigkeiten mit npm ci..."
    Invoke-NativeCommand { npm ci } "npm ci"
  }

  if ($Clean) {
    Write-Host "Hinweis: Alte Builds bleiben erhalten; vorhandene Dateien duerfen nur ueberschrieben werden."
  }

  $buildNumberPath = Join-Path $rootDir ".local-build-number"
  $lastBuildNumber = 0
  if (Test-Path -LiteralPath $buildNumberPath) {
    $storedBuildNumber = (Get-Content -LiteralPath $buildNumberPath -Raw).Trim()
    if (-not [int]::TryParse($storedBuildNumber, [ref]$lastBuildNumber) -or $lastBuildNumber -lt 0) {
      throw "Ungueltige lokale Buildnummer in $buildNumberPath."
    }
  }
  $localBuildNumber = $lastBuildNumber + 1

  Write-Host "Baue Tailwind CSS..."
  Invoke-NativeCommand { npm run build:css } "CSS-Build"

  $env:PUBLISH_MODE = "never"
  $env:LOCAL_BUILD_NUMBER = $localBuildNumber
  Write-Host "Erzeuge lokalen Windows-Build Nr. $localBuildNumber..."
  Invoke-NativeCommand { node ".\scripts\release-with-icon.js" } "Windows-Build"
  Set-Content -LiteralPath $buildNumberPath -Value $localBuildNumber -NoNewline

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

  if ($null -eq $previousLocalBuildNumber) {
    Remove-Item Env:LOCAL_BUILD_NUMBER -ErrorAction SilentlyContinue
  }
  else {
    $env:LOCAL_BUILD_NUMBER = $previousLocalBuildNumber
  }

  Pop-Location
}
