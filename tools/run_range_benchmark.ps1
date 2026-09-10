<#
.SYNOPSIS
Measures the multi-Shape apply: one private transaction against a ShapeRange.

.DESCRIPTION
A transaction carries no target - the receiver is the target - so a transaction
holding several picture fills cannot exist. What can exist is a receiver that
stands for several Shapes, and ShapeRange.Fill is one: the same PPCORE wrapper,
the same OART FillFormat, the same validated receiver layout. One apply fills
every member.

This measures it at 1, 2, 8, 32 and 100 Shapes against the same number of
individual applies, and reports the gate - classifying every member, which is
what stops a Connector in the range from reaching the private backend -
separately from the apply, because the gate is the part that grows with N.
#>
param([int]$Iterations = 200, [int]$Runs = 3)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'
if (-not (Test-Path $texture)) { throw "Missing $texture - run tools/generate_textures.ps1" }

function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return [double]$bits[1] }
    }
    return 0.0
}
function Median($values) {
    $sorted = @($values | Sort-Object)
    return $sorted[[int]($sorted.Count / 2)]
}

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$rows = @()
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    $handle = $engine.LoadTexture($bytes)

    foreach ($count in 1, 2, 8, 32, 100) {
        $names = @()
        for ($i = 0; $i -lt $count; $i++) {
            $shape = $slide.Shapes.AddShape(1,
                10 + ($i % 10) * 30, 10 + [math]::Floor($i / 10) * 30, 24, 24)
            $shape.Name = "Bb$count`_$i"
            $names += $shape.Name
        }
        $range = $slide.Shapes.Range($names)

        # Both legs run inside the experiment, in process. Driving the one-at-a-
        # time leg from here would add a cross-process round trip per Shape -
        # about 5.9 ms against the 0.19 being compared - and would flatter the
        # range by twenty-five times for reasons that are not Office's.
        $rangeTotals = @(); $rangeApplies = @(); $rangeGates = @(); $loopTotals = @()
        for ($run = 1; $run -le $Runs; $run++) {
            $report = $engine.ApplyTextureToRange($range, $handle, $Iterations)
            $rangeTotals += Get-Field $report 'totalMeanMs'
            $rangeApplies += Get-Field $report 'applyMeanMs'
            $rangeGates += Get-Field $report 'gateMeanMs'
            $loopTotals += Get-Field $report 'loopMeanMs'
        }

        $filled = ($names | ForEach-Object { $slide.Shapes.Item($_).Fill.Type } |
            Where-Object { $_ -eq 6 }).Count
        $rangeTotal = Median $rangeTotals
        $loopTotal = Median $loopTotals
        $rows += [pscustomobject]@{
            Shapes        = $count
            RangeTotalMs  = [math]::Round($rangeTotal, 4)
            RangeApplyMs  = [math]::Round((Median $rangeApplies), 4)
            GateMs        = [math]::Round((Median $rangeGates), 4)
            PerShapeMs    = [math]::Round($rangeTotal / $count, 5)
            LoopTotalMs   = [math]::Round($loopTotal, 4)
            Speedup       = "{0:N1}x" -f $(if ($rangeTotal -gt 0) { $loopTotal / $rangeTotal } else { 0 })
            Filled        = "$filled/$count"
        }
        foreach ($name in $names) { $slide.Shapes.Item($name).Delete() }
    }
    $engine.ReleaseTexture($handle)
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

"iterations: $Iterations per run, runs: $Runs (median reported)"
$rows | Format-Table -AutoSize
'RangeApplyMs is the private apply alone; GateMs is classifying every member.'
'LoopTotalMs is the same fills one Shape at a time, in process, gate included.'
$rows | Export-Csv "$root/artifacts/range_benchmark.csv" -NoTypeInformation
