<#
.SYNOPSIS
  Builds Release/x64 and assembles a self-contained dist\ folder for distribution.
#>
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$bin  = Join-Path $root "src\LiveWebRegion\bin\x64\Release"
$dist = Join-Path $root "dist\LiveWebRegion"

# 1) Build
& (Join-Path $PSScriptRoot "build.ps1") -NoRegister

# 2) Clean dist
if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Force -Path $dist | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "runtimes\win-x64\native") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "assets") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $dist "samples") | Out-Null

# 3) Copy binaries (WebView2Loader must keep its runtimes\win-x64\native path)
Copy-Item (Join-Path $bin "LiveWebRegion.dll") $dist
Copy-Item (Join-Path $bin "Microsoft.Web.WebView2.Core.dll") $dist
Copy-Item (Join-Path $bin "Microsoft.Web.WebView2.WinForms.dll") $dist
Copy-Item (Join-Path $bin "runtimes\win-x64\native\WebView2Loader.dll") (Join-Path $dist "runtimes\win-x64\native")
Copy-Item (Join-Path $bin "assets\*") (Join-Path $dist "assets")

# 4) Installer scripts + sample + license + check
Copy-Item (Join-Path $PSScriptRoot "register.ps1") $dist
Copy-Item (Join-Path $PSScriptRoot "unregister.ps1") $dist
Copy-Item (Join-Path $PSScriptRoot "webview2-check.ps1") $dist
Copy-Item (Join-Path $root "LICENSE.txt") $dist
Copy-Item (Join-Path $root "samples\demo.html") (Join-Path $dist "samples")

# 5) install.ps1 / uninstall.ps1 (end-user entry points)
@'
# Installs the Live Web Region PowerPoint add-in for the current user.
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "webview2-check.ps1")
if (-not (Test-WebView2Runtime)) {
    Write-Host "WebView2-Runtime fehlt. Bitte zuerst installieren:" -ForegroundColor Yellow
    Write-Host "  https://developer.microsoft.com/microsoft-edge/webview2/  (Evergreen Standalone)" -ForegroundColor Yellow
    Start-Process "https://developer.microsoft.com/microsoft-edge/webview2/"
    Read-Host "Enter druecken, sobald die Runtime installiert ist (oder Strg+C zum Abbrechen)"
}
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning "Bitte PowerPoint schliessen und install.ps1 erneut ausfuehren."
    return
}
& (Join-Path $PSScriptRoot "register.ps1") -DllPath (Join-Path $PSScriptRoot "LiveWebRegion.dll")
Write-Host "Fertig. PowerPoint starten -> Reiter 'Live Web'." -ForegroundColor Green
'@ | Set-Content -Path (Join-Path $dist "install.ps1") -Encoding UTF8

@'
# Removes the Live Web Region add-in for the current user.
& (Join-Path $PSScriptRoot "unregister.ps1")
'@ | Set-Content -Path (Join-Path $dist "uninstall.ps1") -Encoding UTF8

# Double-click entry points (.cmd) so end users don't need to know PowerShell.
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

Zeigt eine lokale HTML/JS/CSS-Datei live in einem Folienbereich an - auch
interaktiv im Praesentationsmodus.

INSTALLATION (pro Benutzer, ohne Admin)
  1. Rechtsklick auf install.ps1 -> "Mit PowerShell ausfuehren"
     (oder in PowerShell:  powershell -ExecutionPolicy Bypass -File .\install.ps1)
  2. Falls die WebView2-Runtime fehlt, wird der Download geoeffnet.
  3. PowerPoint starten -> Reiter "Live Web".

VERWENDEN
  1. Eine Form (z. B. Rechteck) auf der Folie einfuegen und markieren.
  2. Reiter "Live Web" -> "Bereich festlegen" -> HTML-Datei waehlen.
  3. Praesentation starten (F5): die Seite laeuft live im Bereich.
     - Navigation: Pfeiltasten / Bild auf-ab, Esc beendet - auch wenn die
       Seite gerade den Fokus hat. Klicks/Tippen gehen an die Seite.

DEINSTALLATION
  uninstall.ps1 ausfuehren.

VORAUSSETZUNGEN
  Windows, PowerPoint Desktop x64, Microsoft Edge WebView2 Runtime.
'@ | Set-Content -Path (Join-Path $dist "README.txt") -Encoding UTF8

Write-Host "`nPaket erstellt unter: $dist" -ForegroundColor Green
Get-ChildItem $dist -Recurse -File | ForEach-Object { Write-Host ("  " + $_.FullName.Substring($dist.Length + 1)) }
