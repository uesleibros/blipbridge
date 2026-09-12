<#
.SYNOPSIS
Dumps the structure behind Shape.Fill in whichever PowerPoint is running.

.DESCRIPTION
The first step in re-deriving the receiver walk for a new architecture, and it is
deliberately a dump rather than a decoder: it reports the wrapper object's
leading words and its vtable slots with raw bytes, and draws no conclusions. What
a thunk means is for a human with a disassembler to say.

Nothing here calls into Office internals. Every access is a guarded read.

Uses a disposable presentation and a disposable Shape. The user's own content is
never involved in reconnaissance.

.PARAMETER Slots
How many vtable slots to dump. Default 40.

.PARAMETER Bytes
How many bytes to show at each slot target. Default 16.

.PARAMETER Output
Where to write the dump. Defaults to artifacts/fill_structure_<arch>.txt.
#>
param([int]$Slots = 40, [int]$Bytes = 16, [string]$Output)

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

$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$process = Get-Process POWERPNT | Select-Object -First 1
$architecture = 'x64'
try {
    $signature = @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool IsWow64Process(IntPtr process, out bool result);
'@
    $helper = Add-Type -MemberDefinition $signature -Name BbWow2 -Namespace BbProbe2 -PassThru
    $isWow = $false
    if ($helper::IsWow64Process($process.Handle, [ref]$isWow) -and $isWow) { $architecture = 'x86' }
} catch { }

if (-not $Output) { $Output = Join-Path $root "artifacts/fill_structure_$architecture.txt" }

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, 40, 40, 120, 120)

    # A plain AutoShape with no picture fill yet: the simplest object the walk
    # starts from, and the one the accelerated backend already understands on x64.
    $dump = $engine.InspectFillStructure($shape.Fill, $Slots, $Bytes)

    $header = @(
        "Shape.Fill structure dump",
        "architecture   : $architecture",
        "POWERPNT       : $((Get-Item $process.Path).VersionInfo.FileVersion)",
        "path           : $($process.Path)",
        "recorded       : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "",
        "Read-only. Nothing below was called.",
        ""
    ) -join "`r`n"

    New-Item -ItemType Directory -Path (Split-Path $Output -Parent) -Force | Out-Null
    ($header + $dump) | Set-Content $Output -Encoding utf8
    Write-Host $header
    Write-Host $dump
    Write-Host "written to $Output"
} finally {
    $presentation.Saved = $true
    $presentation.Close()
    $addin.Connect = $false
}
