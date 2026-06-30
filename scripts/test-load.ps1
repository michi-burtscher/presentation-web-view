<#
.SYNOPSIS
  Clean load test: kill PowerPoint, clear diagnostics, real-launch with a blank
  presentation, then report whether the add-in connected.
#>
param(
    [int]$WaitSeconds = 12
)

$exe   = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$blank = Join-Path $PSScriptRoot "..\samples\blank.pptx"
$diag  = "HKCU:\Software\LiveWebRegion\Diag"
$logLocal = Join-Path $env:LOCALAPPDATA "LiveWebRegion\addin.log"
$logTemp  = Join-Path $env:TEMP "LiveWebRegion\addin.log"

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2

# Office hard-disables an add-in after a prior crash by adding it to DisabledItems
# and re-applies that on every startup until cleared. Clear ours each run.
foreach ($ver in @("16.0","15.0")) {
    $di = "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Resiliency\DisabledItems"
    if (Test-Path $di) {
        foreach ($name in (Get-Item $di).Property) {
            $bytes = (Get-ItemProperty $di).$name
            $txt = (-join ($bytes | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ })).ToLower()
            if ($txt -like "*livewebregion*") {
                Remove-ItemProperty -Path $di -Name $name -Force
                Write-Host "  (cleared stale DisabledItems entry '$name')" -ForegroundColor DarkYellow
            }
        }
    }
}

Remove-Item $diag -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $logLocal -Force -ErrorAction SilentlyContinue
Remove-Item $logTemp  -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $blank)) {
    Write-Host "Erzeuge blank.pptx ..." -ForegroundColor DarkGray
    $a = New-Object -ComObject PowerPoint.Application
    $p = $a.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
    [void]$p.Slides.Add(1, 12)
    $p.SaveAs((Resolve-Path (Split-Path $blank)).Path + "\blank.pptx")
    $p.Close(); $a.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null
    Start-Sleep 2
}

Write-Host "Starte PowerPoint (echt) mit blank.pptx ..." -ForegroundColor Cyan
Start-Process -FilePath $exe -ArgumentList $blank
Start-Sleep $WaitSeconds

$d = Get-ItemProperty $diag -ErrorAction SilentlyContinue
if ($d) {
    Write-Host "BEACONS:" -ForegroundColor Green
    $d.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { Write-Host ("  {0} = {1}" -f $_.Name, $_.Value) }
} else {
    Write-Host "KEINE Beacons - OnConnection feuerte nicht." -ForegroundColor Yellow
}

$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue
Write-Host ("PowerPoint laeuft noch: " + [bool]$pp)
if ($pp) {
    $m = $pp.Modules | Where-Object { $_.ModuleName -like 'LiveWebRegion*' }
    Write-Host ("LiveWebRegion.dll geladen: " + [bool]$m)
}
foreach ($lp in @($logLocal, $logTemp)) {
    if (Test-Path $lp) { Write-Host "=== LOG @ $lp ===" -ForegroundColor Green; Get-Content $lp }
}

# Surface any fresh .NET crash
$evt = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddMinutes(-2); ProviderName='.NET Runtime'} -ErrorAction SilentlyContinue | Select-Object -First 1
if ($evt) { Write-Host "!! .NET CRASH:" -ForegroundColor Red; Write-Host (($evt.Message -split "`n")[0..2] -join " ") }
