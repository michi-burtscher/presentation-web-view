<#
.SYNOPSIS
  Builds the add-in (x64, Release) and re-registers it.
.NOTES
  Close PowerPoint before running, otherwise the DLL is locked.
#>
param([switch]$NoRegister)

$ErrorActionPreference = "Stop"
$proj = Join-Path $PSScriptRoot "..\src\LiveWebRegion\LiveWebRegion.csproj"

# Make sure dotnet is reachable even if not on PATH in this shell.
$dotnet = "dotnet"
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    $dotnet = Join-Path $env:ProgramFiles "dotnet\dotnet.exe"
}

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    Write-Warning "PowerPoint is running - the DLL may be locked. Close it and re-run."
}

& $dotnet build $proj -c Release -p:Platform=x64
if ($LASTEXITCODE -ne 0) { throw "Build failed." }

if (-not $NoRegister) {
    & (Join-Path $PSScriptRoot "register.ps1")
}
