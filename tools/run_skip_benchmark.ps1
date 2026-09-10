<#
.SYNOPSIS
Measures BB_ApplyTextureIfChanged against BB_ApplyTexture, six ways.

.DESCRIPTION
The apply profile put 76% of a cached apply inside one private OART call - the
document edit. That cost cannot be reduced from outside Office, so the only way
past it is to not make the call. This asks what skipping is really worth once
the cost of proving the skip is safe is counted, and what it charges a caller
who never gets to skip.

The legs, in order: repeated ApplyTexture (the baseline), IfChanged on a Shape
that was just invalidated (a genuine first apply), IfChanged repeated (the skip),
two images alternating both ways (the no-gain overhead), delete-and-recreate (a
value-keyed cache's dangerous case), and many Shapes sharing one texture.

Each leg asserts the skip decision it expected, so a run that quietly stopped
skipping fails rather than reporting a spectacular improvement.

Everything runs inside PowerPoint with a windowed presentation, because a
windowless one does not do the same document bookkeeping.
#>
param([int]$Iterations = 5000, [int]$Shapes = 32, [int]$Runs = 3)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$reports = @()
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    for ($run = 1; $run -le $Runs; $run++) {
        Write-Host "run $run of $Runs..."
        $reports += , ($engine.BenchmarkApplySkip($slide, $Shapes, $Iterations))
    }
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

function ConvertTo-Map($report) {
    $map = @{}
    foreach ($pair in ($report -split ';' | Where-Object { $_ -ne '' })) {
        $parts = $pair -split '=', 2
        $map[$parts[0]] = $parts[1]
    }
    return $map
}

$maps = $reports | ForEach-Object { ConvertTo-Map $_ }
$reports | Set-Content "$root/artifacts/skip_benchmark.txt"

# Median across runs, per §19: one run's mean is a sample of the machine's mood.
function Median-Of($maps, $key) {
    $values = @($maps | ForEach-Object { [double]$_[$key] } | Sort-Object)
    if ($values.Count -eq 0) { return 0.0 }
    return $values[[int]($values.Count / 2)]
}

$legs = [ordered]@{
    'ApplyTexture repeated (baseline)' = 'baseline'
    'IfChanged, first apply'           = 'ifChangedFirst'
    'IfChanged, repeated (skip)'       = 'ifChangedRepeated'
    'IfChanged, alternating A/B'       = 'alternatingIfChanged'
    'ApplyTexture, alternating A/B'    = 'alternatingPlain'
    'IfChanged, delete + recreate'     = 'deleteRecreate'
}

$rows = foreach ($label in $legs.Keys) {
    $prefix = $legs[$label]
    [pscustomobject]@{
        Leg      = $label
        MeanMs   = [math]::Round((Median-Of $maps "${prefix}MeanMs"), 5)
        MedianMs = [math]::Round((Median-Of $maps "${prefix}MedianMs"), 5)
        P95Ms    = [math]::Round((Median-Of $maps "${prefix}P95Ms"), 5)
        P99Ms    = [math]::Round((Median-Of $maps "${prefix}P99Ms"), 5)
        MinMs    = [math]::Round((Median-Of $maps "${prefix}MinMs"), 5)
        MaxMs    = [math]::Round((Median-Of $maps "${prefix}MaxMs"), 5)
    }
}

"iterations: $Iterations per leg, runs: $Runs, shapes: $Shapes"
$rows | Format-Table -AutoSize

$passRows = foreach ($prefix in @('manyShapesPlain', 'manyShapesIfChanged')) {
    [pscustomobject]@{
        Pass       = if ($prefix -eq 'manyShapesPlain') { "$Shapes x ApplyTexture" } else { "$Shapes x IfChanged (all skip)" }
        TotalMs    = [math]::Round((Median-Of $maps "${prefix}MeanMs"), 5)
        PerShapeMs = [math]::Round((Median-Of $maps ($prefix -replace 'manyShapes(.+)', 'manyPerShape$1Ms')), 5)
    }
}
$passRows | Format-Table -AutoSize

$speedup = Median-Of $maps 'skipSpeedup'
$saved = Median-Of $maps 'skipSavesMs'
$overhead = Median-Of $maps 'ifChangedOverheadMs'
$reused = ($maps | ForEach-Object { [int]$_['reusedShapeIds'] } | Measure-Object -Sum).Sum
$wrong = ($maps | ForEach-Object { [int]$_['skipsAfterRecreate'] } | Measure-Object -Sum).Sum

"skip speedup:            {0:N1}x" -f $speedup
"skip saves per call:     {0:N5} ms" -f $saved
"overhead when no gain:   {0:N5} ms" -f $overhead
"recreated Shapes reusing a deleted Id: $reused"
"skips granted after recreation (must be 0): $wrong"
if ($wrong -ne 0) { throw "A recreated Shape was skipped - the skip cache is unsafe" }
