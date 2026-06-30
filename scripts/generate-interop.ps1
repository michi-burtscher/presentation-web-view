<#
.SYNOPSIS
  Generates a correct interop assembly for the "Microsoft Add-In Designer" type
  library (IDTExtensibility2 + ext_* enums) using the .NET Framework
  TypeLibConverter API. Output: libs\Interop.Extensibility.dll
  (namespace "AddInDesignerObjects").

  Must run under Windows PowerShell 5.1 (.NET Framework), not PowerShell 7+.
#>
$ErrorActionPreference = "Stop"

$olb     = "C:\Program Files\Common Files\DESIGNER\MSADDNDR.OLB"
$outDir  = Join-Path $PSScriptRoot "..\libs"
$outFile = Join-Path $outDir "Interop.Extensibility.dll"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$outFile = (Join-Path (Resolve-Path $outDir) "Interop.Extensibility.dll")

if ($PSVersionTable.PSEdition -eq "Core") { throw "Run under Windows PowerShell 5.1, not pwsh." }
if (-not (Test-Path $olb)) { throw "Type library not found: $olb" }

Add-Type -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;

public enum REGKIND { DEFAULT = 0, REGISTER = 1, NONE = 2 }

public static class TlbNative {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void LoadTypeLibEx(string strTypeLibName, REGKIND regKind,
        [MarshalAs(UnmanagedType.Interface)] out object typeLib);
}

public class ImportSink : ITypeLibImporterNotifySink {
    public Assembly ResolveRef(object typeLib) {
        // The Add-In Designer TLB only references stdole; use its GAC PIA.
        return Assembly.Load("stdole, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a");
    }
    public void ReportEvent(ImporterEventKind eventKind, int eventCode, string eventMsg) {
        Console.WriteLine("  [tlb] " + eventMsg);
    }
}
"@

$typeLib = $null
[TlbNative]::LoadTypeLibEx($olb, [REGKIND]::NONE, [ref]$typeLib)

$converter = New-Object System.Runtime.InteropServices.TypeLibConverter
$sink = New-Object ImportSink
$flags = [System.Runtime.InteropServices.TypeLibImporterFlags]::SafeArrayAsSystemArray

$asm = $converter.ConvertTypeLibToAssembly(
    $typeLib, $outFile, $flags, $sink, $null, $null, "AddInDesignerObjects", $null)

# AssemblyBuilder.Save writes relative to the current directory.
$fileName = [System.IO.Path]::GetFileName($outFile)
Push-Location $outDir
try { $asm.Save($fileName) } finally { Pop-Location }

Write-Host "Generated: $outFile" -ForegroundColor Green
Write-Host "Public types:" -ForegroundColor Cyan
([Reflection.Assembly]::ReflectionOnlyLoadFrom($outFile)).GetExportedTypes() |
    ForEach-Object { Write-Host ("  " + $_.FullName) }
