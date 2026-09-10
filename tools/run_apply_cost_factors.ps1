<#
.SYNOPSIS
Asks what the private apply's 0.169 ms is actually made of.

.DESCRIPTION
The stage profiler put 76% of a cached apply inside the receiver's own handler.
That locates the cost without naming it, and the name decides what to build
next: if most of it is rendering the changed Shape, a batch that changes many
Shapes before one redraw is worth building; if it is document bookkeeping, the
cost is per edit and no arrangement of the slide will help.

So this varies one thing at a time around the same apply - which slide the Shape
is on, whether it is visible, whether it is on screen at all, whether the window
is minimised - and reports what each is worth.
#>
param([int]$Iterations = 2000)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    [void]$presentation.Slides.Add(1, $ppLayoutBlank)
    $report = $engine.MeasureApplyCostFactors($presentation, $Iterations)
} finally {
    # PowerPoint refuses to close a presentation while its window is still
    # settling out of the minimised state this experiment puts it through.
    try { $presentation.Windows.Item(1).WindowState = 3 } catch { }
    $presentation.Saved = $msoTrue
    try { $presentation.Close() } catch { Start-Sleep -Milliseconds 500; $presentation.Close() }
}

$map = @{}
foreach ($pair in ($report -split ';' | Where-Object { $_ -ne '' })) {
    $parts = $pair -split '=', 2
    $map[$parts[0]] = $parts[1]
}
$report -split ';' | Where-Object { $_ -ne '' } | Set-Content "$root/artifacts/apply_cost_factors.txt"

$legs = [ordered]@{
    'Shape on the displayed slide' = 'displayed'
    'Shape on another slide'       = 'otherSlide'
    'Shape hidden'                 = 'hidden'
    'Shape off the slide area'     = 'offSlide'
    'Window minimised'             = 'minimised'
}

$baseline = [double]$map['displayedMeanMs']
$rows = foreach ($label in $legs.Keys) {
    $prefix = $legs[$label]
    $mean = [double]$map["${prefix}MeanMs"]
    [pscustomobject]@{
        Arrangement = $label
        MeanMs      = [math]::Round($mean, 5)
        MedianMs    = [math]::Round([double]$map["${prefix}MedianMs"], 5)
        MinMs       = [math]::Round([double]$map["${prefix}MinMs"], 5)
        VsDisplayed = if ($baseline -gt 0) { "{0:N1}%" -f (100.0 * ($mean - $baseline) / $baseline) } else { '' }
    }
}

"iterations: $Iterations per arrangement"
$rows | Format-Table -AutoSize
