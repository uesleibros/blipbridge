<#
.SYNOPSIS
Runs the in-process texture benchmark and records the result.

.DESCRIPTION
The measurement itself happens inside PowerPoint, so the cross-process COM cost
of driving it does not land in the numbers. This script only creates a hidden
presentation, asks the engine to measure, and writes the report.
#>
param([int]$Iterations = 500)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, 12)
    $report = $engine.BenchmarkNativeTexture($slide, $texture, $Iterations)
    $lines = @("iterations=$Iterations", $report)
    $lines | Set-Content "$root/artifacts/texture_benchmark.txt"
    $report -split ';' | Where-Object { $_ }
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}
