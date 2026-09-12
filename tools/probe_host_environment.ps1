<#
.SYNOPSIS
Records exactly which PowerPoint is running and what BlipBridge reports inside it.

.DESCRIPTION
Every claim about a backend is a claim about a host, and "PowerPoint 16.0" is not
a host - a build and an architecture are. This writes both down, from the running
process rather than from a file path, so a result can be attributed later without
anyone having to remember which Office was installed that week.

The architecture is read from the **live POWERPNT.EXE process**, not inferred
from the install directory. `Program Files (x86)` is a strong hint and not
evidence; the loaded module list is.

Writes artifacts/host_environment.txt and prints the same to the console.
#>
param([string]$Output)

$ErrorActionPreference = 'Stop'

function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        try {
            $candidate = New-Object -ComObject PowerPoint.Application
            $null = $candidate.Version
            $null = $candidate.Presentations.Count
            return $candidate
        } catch {
            if ($attempt -eq 15) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

function Get-PeMachine([string]$path) {
    try {
        $stream = [IO.File]::OpenRead($path)
        try {
            $buffer = New-Object byte[] 0x400
            $null = $stream.Read($buffer, 0, $buffer.Length)
        } finally { $stream.Dispose() }
        $peOffset = [BitConverter]::ToInt32($buffer, 0x3C)
        $machine = [BitConverter]::ToUInt16($buffer, $peOffset + 4)
        switch ($machine) {
            0x14C { 'x86' }
            0x8664 { 'x64' }
            default { "0x$('{0:X}' -f $machine)" }
        }
    } catch { "unreadable: $($_.Exception.Message)" }
}

$root = Split-Path $PSScriptRoot -Parent
if (-not $Output) { $Output = Join-Path $root 'artifacts/host_environment.txt' }

$lines = New-Object System.Collections.Generic.List[string]
function Emit([string]$text) { $script:lines.Add($text); Write-Host $text }

$app = Connect-PowerPoint
$process = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $process) { throw 'PowerPoint is reachable over COM but has no POWERPNT process' }

$exe = $process.Path
Emit "PowerPoint product     : $($app.Name)"
Emit "Application.Version    : $($app.Version)"
Emit "POWERPNT.EXE path      : $exe"
Emit "POWERPNT.EXE version   : $((Get-Item $exe).VersionInfo.FileVersion)"
Emit "POWERPNT.EXE machine   : $(Get-PeMachine $exe)"

# The decisive test: a 64-bit process cannot load a 32-bit module and vice
# versa, so the modules the live process has loaded settle its architecture.
$wow = $null
try {
    $handle = $process.Handle
    $signature = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool IsWow64Process(IntPtr process, out bool result);
'@
    $helper = Add-Type -MemberDefinition $signature -Name BbWow -Namespace BbProbe -PassThru
    $isWow = $false
    if ($helper::IsWow64Process($handle, [ref]$isWow)) { $wow = $isWow }
} catch { }
if ($null -ne $wow) {
    Emit "process is WOW64       : $wow  (true means a 32-bit process on 64-bit Windows)"
}

Emit ''
Emit 'Office modules loaded by the running process:'

<#
A 64-bit process enumerating a WOW64 process sees only the 64-bit WOW layer -
ntdll and wow64*.dll, seven modules - and none of Office. The 32-bit modules are
visible only to a 32-bit reader, so when the target is WOW64 the enumeration is
delegated to 32-bit PowerShell. Reporting the seven WOW modules as "the Office
modules" would look like an answer and be nothing of the kind.
#>
$moduleBody = @'
param($processId)
$target = Get-Process -Id $processId
foreach ($module in $target.Modules) {
    if ($module.ModuleName -match '^(POWERPNT\.EXE|OART\.DLL|GFX\.DLL|PPCORE\.DLL|MSO\.DLL|MSO[0-9]+WIN32CLIENT\.DLL|VBE7\.DLL|BlipBridge.*\.DLL)$') {
        '{0}|{1}' -f $module.ModuleName, $module.FileVersionInfo.FileVersion
    }
}
'@

$rows = @()
$helperPath = Join-Path ([IO.Path]::GetTempPath()) ('bb_modules_' + [guid]::NewGuid().ToString('N') + '.ps1')
Set-Content $helperPath $moduleBody -Encoding utf8
try {
    $shell = if ($wow -eq $true) {
        Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    } else {
        Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }
    $rows = & $shell -NoProfile -ExecutionPolicy Bypass -File $helperPath $process.Id
} finally { Remove-Item $helperPath -ErrorAction SilentlyContinue }

if (-not $rows) {
    Emit '  (none - the Office modules load lazily; open a presentation and fill a Shape first)'
} else {
    foreach ($row in $rows) {
        $parts = "$row" -split '\|', 2
        if ($parts.Length -eq 2) { Emit ("  {0,-28} {1}" -f $parts[0], $parts[1]) }
    }
}

Emit ''
$configuration = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if (Test-Path $configuration) {
    $c = Get-ItemProperty $configuration
    Emit "Click-to-Run platform  : $($c.Platform)"
    Emit "Click-to-Run version   : $($c.ClientVersionToReport)"
    Emit "Click-to-Run products  : $($c.ProductReleaseIds)"
}

Emit ''
Emit "harness PowerShell     : $(if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' })"
Emit "Windows                : $([Environment]::OSVersion.VersionString)"
Emit "recorded               : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

# What BlipBridge itself reports in this host.
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
try {
    $engine = $addin.Object
    $report = $engine.AbiCapabilities()
    $mask = 0
    $backend = ''
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -ne 2) { continue }
        if ($bits[0] -eq 'mask') { $mask = [uint32]$bits[1] }
        if ($bits[0] -eq 'backend') { $backend = $bits[1] }
    }
    Emit ''
    Emit "BlipBridge backend     : $backend"
    Emit ("BlipBridge mask        : 0x{0:X4}" -f $mask)
    Emit ("IsAvailable            : {0}   (mask is non-zero)" -f ($mask -ne 0))
    Emit ("IsAccelerated          : {0}   (BB_CAP_NATIVE_BACKEND)" -f (($mask -band 0x1) -ne 0))
} finally {
    $addin.Connect = $false
}

New-Item -ItemType Directory -Path (Split-Path $Output -Parent) -Force | Out-Null
$lines | Set-Content $Output -Encoding utf8
Write-Host ''
Write-Host "written to $Output"
