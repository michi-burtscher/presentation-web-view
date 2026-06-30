<#
.SYNOPSIS
  Removes the LiveWebRegion COM add-in registration for the current user.
#>
$ErrorActionPreference = "SilentlyContinue"

$Clsid  = "{7E9B2C14-3A6D-4F58-9C1E-2B7A5D0F8E31}"
$ProgId = "LiveWebRegion.AddIn"

Remove-Item -Path "HKCU:\Software\Classes\CLSID\$Clsid" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\$ProgId" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$ProgId" -Recurse -Force

Write-Host "Unregistered '$ProgId'." -ForegroundColor Green
