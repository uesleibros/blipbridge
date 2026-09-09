<#
.SYNOPSIS
Breaks the picture-fill cost down stage by stage.

.DESCRIPTION
Applying an existing texture costs ~0.19 ms; creating new content and applying it
costs ~0.9-1.0 ms. This attributes that difference to named stages instead of
guessing at it, across four legs:

  reuse        one texture applied repeatedly
  newSame      the same encoded bytes re-created every iteration
  newDistinct  raw pixels whose content genuinely changes every iteration
  pool         a ring of eight pre-created distinct textures, applied in turn

The presentation is created **with a window**, because a windowless presentation
does not do the same document bookkeeping - the same trap that made the first
undo harness draw a wrong conclusion. Everything runs in process; a PowerShell
loop would add a cross-process round trip larger than the stages being measured.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

$msoTrue = -1
$ppLayoutBlank = 12
$image = Join-Path $root 'artifacts/textures/texture_128_0.png'
if (-not (Test-Path $image)) {
    throw "Missing profiler image $image - run tools/generate_textures.ps1 first"
}

$iterations = 300
if ($args.Count -ge 1) { $iterations = [int]$args[0] }

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $report = $engine.ProfileFillStages($slide, $image, $iterations)
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

# One key=value per line: the report is long enough that a single line is
# unreadable, and every consumer of it is either a human or a grep.
$lines = $report -split ';' | Where-Object { $_ -ne '' }
$lines | Set-Content "$root/artifacts/stage_profile.txt"
$lines
