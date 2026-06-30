<#
.SYNOPSIS
  Builds Release/x64, assembles a self-contained dist\ folder, and builds a
  single self-contained installer EXE (LiveWebRegionSetup.exe).
#>
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$bin  = Join-Path $root "src\LiveWebRegion\bin\x64\Release"
$dist = Join-Path $root "dist\LiveWebRegion"

$dotnet = "dotnet"
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) { $dotnet = Join-Path $env:ProgramFiles "dotnet\dotnet.exe" }

# 1) Build the add-in
& (Join-Path $PSScriptRoot "build.ps1") -NoRegister

# 2) Stage the app payload (only the files the add-in needs at runtime)
$stage = Join-Path $env:TEMP "LiveWebRegion_stage"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path (Join-Path $stage "runtimes\win-x64\native") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $stage "assets") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $stage "samples") | Out-Null
Copy-Item (Join-Path $bin "LiveWebRegion.dll") $stage
Copy-Item (Join-Path $bin "Microsoft.Web.WebView2.Core.dll") $stage
Copy-Item (Join-Path $bin "Microsoft.Web.WebView2.WinForms.dll") $stage
Copy-Item (Join-Path $bin "runtimes\win-x64\native\WebView2Loader.dll") (Join-Path $stage "runtimes\win-x64\native")
Copy-Item (Join-Path $bin "assets\*") (Join-Path $stage "assets")
Copy-Item (Join-Path $root "samples\demo.html") (Join-Path $stage "samples")

# 3) Zip the payload and build the installer EXE that embeds it
$payload = Join-Path $root "src\Installer\payload.zip"
if (Test-Path $payload) { Remove-Item $payload -Force }
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $payload -Force
& $dotnet build (Join-Path $root "src\Installer\Installer.csproj") -c Release | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Installer build failed." }
$setupExe = Join-Path $root "src\Installer\bin\Release\LiveWebRegionSetup.exe"

# 4) Assemble dist (manual install variant: files + scripts)
if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Force -Path (Join-Path $dist "runtimes\win-x64\native") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "assets") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "samples") | Out-Null
Copy-Item (Join-Path $stage "*") $dist -Recurse -Force
Copy-Item (Join-Path $PSScriptRoot "register.ps1") $dist
Copy-Item (Join-Path $PSScriptRoot "unregister.ps1") $dist
Copy-Item (Join-Path $PSScriptRoot "webview2-check.ps1") $dist
Copy-Item (Join-Path $root "LICENSE.txt") $dist
Copy-Item $setupExe $dist   # the one-click EXE installer

# 5) install/uninstall ps1 (used by the .cmd wrappers)
@'
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "webview2-check.ps1")
if (-not (Test-WebView2Runtime)) {
    Write-Host "WebView2-Runtime fehlt. Download wird geoeffnet:" -ForegroundColor Yellow
    Start-Process "https://developer.microsoft.com/microsoft-edge/webview2/"
    Read-Host "Enter druecken, sobald die Runtime installiert ist (oder Strg+C zum Abbrechen)"
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning "Bitte PowerPoint schliessen und install.ps1 erneut ausfuehren."; return
}
& (Join-Path $PSScriptRoot "register.ps1") -DllPath (Join-Path $PSScriptRoot "LiveWebRegion.dll")
Write-Host "Fertig. PowerPoint starten -> Reiter 'Live Web'." -ForegroundColor Green
'@ | Set-Content -Path (Join-Path $dist "install.ps1") -Encoding UTF8

@'
& (Join-Path $PSScriptRoot "unregister.ps1")
'@ | Set-Content -Path (Join-Path $dist "uninstall.ps1") -Encoding UTF8

@'
@echo off
echo Installiere "Live Web Region" PowerPoint Add-in...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
pause
'@ | Set-Content -Path (Join-Path $dist "Install.cmd") -Encoding ASCII

@'
@echo off
echo Entferne "Live Web Region" PowerPoint Add-in...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
echo.
pause
'@ | Set-Content -Path (Join-Path $dist "Uninstall.cmd") -Encoding ASCII

# 6) README
@'
Live Web Region - PowerPoint Add-in
===================================

Zeigt eine Website (lokale HTML-Datei ODER Online-URL) live in einem
Folienbereich an - auch interaktiv im Praesentationsmodus.

INSTALLATION - Variante A (am einfachsten):
  LiveWebRegionSetup.exe  doppelklicken.
  (Self-contained Installer, pro Benutzer, ohne Admin. Deinstallation:
   LiveWebRegionSetup.exe /uninstall)

INSTALLATION - Variante B (Skript):
  Install.cmd  doppelklicken.

In beiden Faellen: fehlt die WebView2-Runtime, wird der Download geoeffnet.
Danach PowerPoint starten -> Reiter "Live Web".

VERWENDEN
  1. Optional eine Form markieren (sonst wird ein Rechteck eingefuegt).
  2. "Link setzen" -> URL (https://...) ODER lokale HTML-Datei.
  3. F5: die Seite laeuft live im Bereich.
     Navigation per Pfeiltasten / Bild auf-ab / Esc - auch wenn die Seite
     den Fokus hat. Klicks/Tippen gehen an die Seite.

DEINSTALLATION
  LiveWebRegionSetup.exe /uninstall   oder   Uninstall.cmd

VORAUSSETZUNGEN
  Windows, PowerPoint Desktop x64, Microsoft Edge WebView2 Runtime.
'@ | Set-Content -Path (Join-Path $dist "README.txt") -Encoding UTF8

Write-Host "`nPaket erstellt unter: $dist" -ForegroundColor Green
Get-ChildItem $dist -Recurse -File | ForEach-Object { Write-Host ("  " + $_.FullName.Substring($dist.Length + 1)) }
