<#
.SYNOPSIS
Measures BB_ApplyTextureBatch against N individual BB_ApplyTexture calls.

.DESCRIPTION
Both legs run in process through the real C ABI, so the numbers isolate the
ABI-side saving: one entry, one thread check and one validation instead of N.
A VBA caller saves more on top of this, because each Declare call also costs an
interpreter-to-native transition - that extra saving is not measured here and is
not claimed.
#>
param([int[]]$ShapeCounts = @(10, 50, 100, 200), [int]$Iterations = 50)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$results = New-Object System.Collections.Generic.List[string]

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, 12)
    foreach ($count in $ShapeCounts) {
        $report = $engine.BenchmarkTextureBatch($slide, $count, $Iterations)
        $results.Add($report)
    }
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

$results | Set-Content "$root/artifacts/batch_benchmark.txt"
$results
