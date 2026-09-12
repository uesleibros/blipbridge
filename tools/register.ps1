<#
.SYNOPSIS
Registers the research COM surface per-user, in the registry view the matching
PowerPoint will actually read.

.DESCRIPTION
The research add-in is an in-process COM server, so it has to be registered in
the same bitness as the host that loads it. That is not a detail a caller can be
left to get right by hand: a 32-bit PowerPoint reads
`HKCU\Software\Classes\Wow6432Node\CLSID`, a 64-bit one reads
`HKCU\Software\Classes\CLSID`, and registering into the wrong one produces
"class not registered" with nothing to suggest why.

Windows does that redirection automatically **for the process doing the writing**,
so the reliable way to register a 32-bit server is to write from a 32-bit
process. This script re-launches itself in 32-bit PowerShell when it is asked to
register a 32-bit DLL and is not already running as one.

.PARAMETER Configuration
Release or Debug, used only to locate the default build tree.

.PARAMETER Architecture
x64 or x86. Decides which DLL is looked for by default and which registry view is
written. Defaults to the architecture of the DLL being registered when -Dll is
given, and to x64 otherwise.

.PARAMETER Dll
Registers a research DLL from somewhere other than the default build tree. The
reason it exists: a build configured with BB_FORCE_PORTABLE_BACKEND puts the
portable backend in the same research surface, and registering that one is how
the Office harnesses exercise the portable backend in a real PowerPoint instead
of only compiling it.

.PARAMETER ResearchAddin
Also writes the PowerPoint add-in key, so the harnesses can connect it explicitly.
#>
param([ValidateSet('Release', 'Debug')][string]$Configuration = 'Release',
      [ValidateSet('x64', 'x86')][string]$Architecture,
      [string]$Dll,
      [switch]$ResearchAddin)

$ErrorActionPreference = 'Stop'

if ($Dll) {
    $dll = (Resolve-Path $Dll).Path
    if (-not $Architecture) {
        # Read the architecture out of the PE header rather than the file name.
        # A name is a convention; the machine field is what the loader uses, and
        # trusting the name is how a 32-bit DLL ends up registered for a 64-bit
        # host and fails at activation instead of here.
        $bytes = [IO.File]::ReadAllBytes($dll)
        $peOffset = [BitConverter]::ToInt32($bytes, 0x3C)
        $machine = [BitConverter]::ToUInt16($bytes, $peOffset + 4)
        $Architecture = if ($machine -eq 0x14C) { 'x86' } elseif ($machine -eq 0x8664) { 'x64' }
                        else { throw "Unrecognised PE machine 0x$('{0:X}' -f $machine) in $dll" }
    }
} else {
    if (-not $Architecture) { $Architecture = 'x64' }
    $dll = (Resolve-Path "$PSScriptRoot/../build/$Configuration/BlipBridgeResearch-$Architecture.dll").Path
}

<#
Re-launch in the matching bitness if we are in the wrong one. Only 32-bit needs
it in practice - this script is normally run from 64-bit PowerShell - but the
check is written both ways so neither direction can silently write the wrong
view.
#>
$needs32 = $Architecture -eq 'x86'
if ($needs32 -ne (-not [Environment]::Is64BitProcess)) {
    $shell = if ($needs32) { "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" }
             else { "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" }
    if (-not (Test-Path $shell)) { throw "Cannot find $Architecture PowerShell at $shell" }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath,
                   '-Architecture', $Architecture, '-Dll', $dll)
    if ($ResearchAddin) { $arguments += '-ResearchAddin' }
    & $shell @arguments
    if ($LASTEXITCODE -ne 0) { throw "Re-launched registration failed with $LASTEXITCODE" }
    return
}

$id = '{2E2E2731-C523-486B-89CB-2A89484F1E32}'
$base = 'HKCU:\Software\Classes'

New-Item "$base/CLSID/$id/InprocServer32" -Force | Out-Null
Set-Item "$base/CLSID/$id/InprocServer32" -Value $dll
New-ItemProperty "$base/CLSID/$id/InprocServer32" -Name ThreadingModel -Value Apartment -Force | Out-Null
New-Item "$base/BlipBridge.Engine/CLSID" -Force | Out-Null
Set-Item "$base/BlipBridge.Engine/CLSID" -Value $id
New-Item "$base/CLSID/$id/ProgID" -Force | Out-Null
Set-Item "$base/CLSID/$id/ProgID" -Value BlipBridge.Engine

if ($ResearchAddin) {
    # The add-in key is not redirected - Office reads it from the same place in
    # both bitnesses - so writing it twice is harmless and writing it once is
    # enough.
    $key = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\BlipBridge.Engine'
    New-Item $key -Force | Out-Null
    New-ItemProperty $key -Name FriendlyName -Value 'BlipBridge research harness' -Force | Out-Null
    New-ItemProperty $key -Name Description -Value 'Explicitly connected native benchmark harness' -Force | Out-Null
    New-ItemProperty $key -Name LoadBehavior -PropertyType DWord -Value 0 -Force | Out-Null
}

$view = if ([Environment]::Is64BitProcess) { '64-bit view' } else { '32-bit (WOW6432Node) view' }
Write-Output "Registered per-user $Architecture in the $view : $dll"
