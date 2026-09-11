<#
.SYNOPSIS
Compares the three ways to fill N Shapes with one texture.

.DESCRIPTION
All three go through the public C ABI, inside PowerPoint's process:

    N x BB_ApplyTexture     one apply per Shape
    BB_ApplyTextureBatch    one ABI crossing, still N applies
    BB_ApplyTextureRange    one apply through Office's own range receiver

The comparison legs have to run in process. Driven from PowerShell they read
about 5.9 ms per Shape - a cross-process Automation round trip rather than
Office - which would make the range look twenty-five times better than it is.

Absolute numbers move by a factor of two or more with what else the machine is
doing, so the speed-up columns are what carry the claim: every leg in a row is
measured in the same run, on the same Shapes, with the same image.
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

    foreach ($count in 1, 2, 8, 16, 32, 64, 100) {
        $names = @()
        for ($i = 0; $i -lt $count; $i++) {
            $shape = $slide.Shapes.AddShape(1,
                10 + ($i % 10) * 30, 10 + [math]::Floor($i / 10) * 30, 24, 24)
            $shape.Name = "Bb$count`_$i"
            $names += $shape.Name
        }
        $range = $slide.Shapes.Range($names)

        # Fewer rounds as N grows: the one-at-a-time leg is N real Office edits
        # per round, and at 100 Shapes that is already 20 ms of work each time.
        $rounds = [math]::Max(20, [int]($Iterations * 8 / [math]::Max(8, $count)))

        $loops = @(); $batches = @(); $ranges = @()
        for ($run = 1; $run -le $Runs; $run++) {
            $report = $engine.ApplyTextureToRange($range, $handle, $rounds)
            $loops += Get-Field $report 'loopMeanMs'
            $batches += Get-Field $report 'batchMeanMs'
            $ranges += Get-Field $report 'rangeMeanMs'
        }

        $filled = ($names | ForEach-Object { $slide.Shapes.Item($_).Fill.Type } |
            Where-Object { $_ -eq 6 }).Count
        $loop = Median $loops
        $batch = Median $batches
        $ranged = Median $ranges
        $rows += [pscustomobject]@{
            Shapes        = $count
            Rounds        = $rounds
            LoopMs        = [math]::Round($loop, 4)
            BatchMs       = [math]::Round($batch, 4)
            RangeMs       = [math]::Round($ranged, 4)
            PerShapeLoop  = [math]::Round($loop / $count, 5)
            PerShapeRange = [math]::Round($ranged / $count, 5)
            VsLoop        = "{0:N1}x" -f $(if ($ranged -gt 0) { $loop / $ranged } else { 0 })
            VsBatch       = "{0:N1}x" -f $(if ($ranged -gt 0) { $batch / $ranged } else { 0 })
            Filled        = "$filled/$count"
        }
        foreach ($name in $names) { $slide.Shapes.Item($name).Delete() }
    }
    $null = $engine.ReleaseTexture($handle)
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

"runs: $Runs (median reported); all three legs through the public C ABI, in process"
$rows | Format-Table -AutoSize
$rows | Export-Csv "$root/artifacts/range_benchmark.csv" -NoTypeInformation
