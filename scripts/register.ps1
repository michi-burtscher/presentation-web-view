<#
.SYNOPSIS
  Registers the LiveWebRegion managed COM add-in for the current user (no admin).

.DESCRIPTION
  Writes the managed-COM CLSID entries (hosted by mscoree.dll) plus the PowerPoint
  Addins key under HKCU. PowerPoint must be closed while running this.
#>
param(
    [string]$DllPath = (Join-Path $PSScriptRoot "..\src\LiveWebRegion\bin\x64\Release\LiveWebRegion.dll")
)

$ErrorActionPreference = "Stop"

$Clsid       = "{7E9B2C14-3A6D-4F58-9C1E-2B7A5D0F8E31}"
$ProgId      = "LiveWebRegion.AddIn"
$AssemblyVer = "1.0.0.0"
$AssemblyFull = "LiveWebRegion, Version=$AssemblyVer, Culture=neutral, PublicKeyToken=null"
$Runtime     = "v4.0.30319"
$Mscoree     = Join-Path $env:WINDIR "System32\mscoree.dll"

$DllPath = (Resolve-Path $DllPath).Path
if (-not (Test-Path $DllPath)) { throw "DLL not found: $DllPath. Build first." }
$CodeBase = ([Uri]$DllPath).AbsoluteUri   # file:///C:/...

Write-Host "Registering '$ProgId'" -ForegroundColor Cyan
Write-Host "  DLL : $DllPath"

function Set-Key([string]$Path, [string]$Name, [string]$Value) {
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    if ($Name -eq "") {
        Set-ItemProperty -Path $Path -Name "(default)" -Value $Value
    } else {
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
    }
}

# --- COM CLSID (managed, hosted by mscoree) ---
$clsidKey = "HKCU:\Software\Classes\CLSID\$Clsid"
Set-Key $clsidKey "" $ProgId

$inproc = "$clsidKey\InprocServer32"
Set-Key $inproc "" $Mscoree
Set-Key $inproc "ThreadingModel" "Both"
Set-Key $inproc "Class"          $ProgId
Set-Key $inproc "Assembly"       $AssemblyFull
Set-Key $inproc "RuntimeVersion" $Runtime
Set-Key $inproc "CodeBase"       $CodeBase

# versioned subkey (some loaders look here)
$inprocVer = "$inproc\$AssemblyVer"
Set-Key $inprocVer "Class"          $ProgId
Set-Key $inprocVer "Assembly"       $AssemblyFull
Set-Key $inprocVer "RuntimeVersion" $Runtime
Set-Key $inprocVer "CodeBase"       $CodeBase

Set-Key "$clsidKey\ProgId" "" $ProgId

# --- ProgId -> CLSID ---
Set-Key "HKCU:\Software\Classes\$ProgId" "" $ProgId
Set-Key "HKCU:\Software\Classes\$ProgId\CLSID" "" $Clsid

# --- PowerPoint add-in registration ---
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$ProgId"
Set-Key $addinKey "FriendlyName" "Live Web Region"
Set-Key $addinKey "Description"  "Zeigt HTML/JS live in einem Folienbereich an."
New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "CommandLineSafe" -Value 0 -PropertyType DWord -Force | Out-Null

# --- Clear any "disabled item" PowerPoint may have set after a crash ---
# These survive a LoadBehavior reset and silently keep the add-in from loading.
foreach ($ver in @("16.0","15.0")) {
    $di = "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Resiliency\DisabledItems"
    if (Test-Path $di) {
        foreach ($name in (Get-Item $di).Property) {
            $bytes = (Get-ItemProperty $di).$name
            $txt = (-join ($bytes | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ })).ToLower()
            if ($txt -like "*livewebregion*") {
                Remove-ItemProperty -Path $di -Name $name -Force
                Write-Host "  Cleared DisabledItems entry '$name'." -ForegroundColor DarkYellow
            }
        }
    }
}

Write-Host "Done. Start PowerPoint and look for the 'Live Web' ribbon tab." -ForegroundColor Green
