<#
.SYNOPSIS
Attributes the cost of a cached ApplyTexture, stage by stage.

.DESCRIPTION
Runs the production apply path in process, thousands of times, timing each step
separately. The point is to find out where ~0.186 ms actually goes before trying
to make it smaller.

Everything runs inside PowerPoint with a windowed presentation, because a
windowless one does not do the same document bookkeeping - the trap that made an
earlier undo harness draw a wrong conclusion.
#>
param([int]$Iterations = 5000)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'
if (-not (Test-Path $texture)) { throw "Missing $texture - run tools/generate_textures.ps1" }

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shape = $slide.Shapes.AddShape(1, 20, 20, 200, 200)
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    $handle = $engine.LoadTexture($bytes)

    $report = $engine.ProfileApplyStages($shape, $handle, $Iterations)

    $engine.ReleaseTexture($handle)
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

$lines = $report -split ';' | Where-Object { $_ -ne '' }
$lines | Set-Content "$root/artifacts/apply_profile.txt"

# A table rather than a wall of key=value: the whole point is to see which stage
# dominates at a glance.
$stages = [ordered]@{}
foreach ($line in $lines) {
    if ($line -match '^apply\.([A-Za-z]+)\.(mean|median|p95|p99|min|max)=(.+)$') {
        $name = $Matches[1]; $metric = $Matches[2]; $value = [double]$Matches[3]
        if (-not $stages.Contains($name)) { $stages[$name] = [ordered]@{} }
        $stages[$name][$metric] = $value
    }
}

$total = 0.0
if ($stages.Contains('total')) { $total = $stages['total']['mean'] }

$rows = foreach ($name in $stages.Keys) {
    $s = $stages[$name]
    $share = 0.0
    if ($total -gt 0 -and $name -ne 'total') { $share = 100.0 * $s['mean'] / $total }
    [pscustomobject]@{
        Stage    = $name
        MeanMs   = [math]::Round($s['mean'], 5)
        MedianMs = [math]::Round($s['median'], 5)
        P95Ms    = [math]::Round($s['p95'], 5)
        P99Ms    = [math]::Round($s['p99'], 5)
        MinMs    = [math]::Round($s['min'], 5)
        MaxMs    = [math]::Round($s['max'], 5)
        Share    = if ($name -eq 'total') { '' } else { "{0:N1}%" -f $share }
    }
}

"iterations: $Iterations"
$rows | Format-Table -AutoSize
$lines | Where-Object { $_ -like 'apply.stageSum=*' -or $_ -like 'apply.unattributed=*' }
