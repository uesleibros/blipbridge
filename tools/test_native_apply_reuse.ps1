<#
.SYNOPSIS
Applies ONE cached image to many Shapes, and probes the reference delta.

.DESCRIPTION
The performance-critical question: can a single GFX cached image be applied to
several Shapes without Office creating a new one each time. The experiment
decodes once and reports that single cached-image address alongside a per-Shape
reference timeline, so reuse is visible rather than asserted.

Coverage: the same Shape twice, two Shapes on one slide, a Shape on another
slide, several AutoShapes, and a Freeform.

It also isolates the open reference-count question from the first native apply,
by applying twice to the same Shape. The first apply replaces a solid fill and
the second replaces a picture fill, so comparing their deltas separates
"establishing a fill" from "replacing one".
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]
$deck = Join-Path $root 'artifacts/native_apply_reuse.pptx'

function Get-Identity($shape) {
    @($shape.Id, $shape.Name, $shape.Type, $shape.Left, $shape.Top, $shape.Width,
      $shape.Height, $shape.Rotation, $shape.ZOrderPosition) -join '|'
}

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $first = $presentation.Slides.Add(1, 12)
    $second = $presentation.Slides.Add(2, 12)

    # Warm GFX/OART on a throwaway Shape. No target Shape ever sees UserPicture.
    $warmup = $first.Shapes.AddShape(1, 600, 400, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    $shapes = @(
        $first.Shapes.AddShape(1, 20, 20, 100, 100),
        $first.Shapes.AddShape(1, 140, 20, 100, 100),
        $first.Shapes.AddShape(5, 260, 20, 100, 100),
        $second.Shapes.AddShape(1, 20, 20, 100, 100)
    )
    $builder = $first.Shapes.BuildFreeform(0, 60, 200)
    $builder.AddNodes(0, 0, 200, 200)
    $builder.AddNodes(0, 0, 130, 340)
    $builder.AddNodes(0, 0, 60, 200)
    $freeform = $builder.ConvertToShape()
    $shapes += $freeform
    for ($i = 0; $i -lt $shapes.Count; $i++) { $shapes[$i].Name = "BB_reuse_$i" }

    $identities = $shapes | ForEach-Object { Get-Identity $_ }
    $nodesBefore = $freeform.Nodes.Count
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)

    # One decode, five Shapes: two on slide 1 as plain AutoShapes, a different
    # AutoShape type, one on slide 2, and a Freeform.
    $fills = @($shapes | ForEach-Object { $_.Fill })
    $report = $engine.NativeApplyReuseExperiment($fills, $bytes)
    $results.Add("reuse report: $report")

    for ($i = 0; $i -lt $shapes.Count; $i++) {
        if ((Get-Identity $shapes[$i]) -ne $identities[$i]) { throw "Shape $i identity changed" }
        if ([int]$shapes[$i].Fill.Type -ne 6) { throw "Shape $i did not get a picture fill" }
    }
    $results.Add("All $($shapes.Count) Shapes kept identity and got a picture fill from one cached image.")

    if ($freeform.Nodes.Count -ne $nodesBefore) { throw 'Freeform node count changed' }
    $freeform.Nodes.SetPosition(2, 220, 200)
    $results.Add("Freeform kept $nodesBefore editable nodes.")

    # Same Shape twice: the second apply replaces a picture fill rather than a
    # solid one, which is what separates the two reference deltas.
    $repeat = $first.Shapes.AddShape(1, 380, 20, 100, 100)
    $repeat.Name = 'BB_reuse_repeat'
    $repeatIdentity = Get-Identity $repeat
    $results.Add("repeat #1 (over solid):  $($engine.NativeApplyExperiment($repeat.Fill, $bytes))")
    $results.Add("repeat #2 (over picture): $($engine.NativeApplyExperiment($repeat.Fill, $bytes))")
    $results.Add("repeat #3 (over picture): $($engine.NativeApplyExperiment($repeat.Fill, $bytes))")
    if ((Get-Identity $repeat) -ne $repeatIdentity) { throw 'Repeated apply changed identity' }
    if ([int]$repeat.Fill.Type -ne 6) { throw 'Repeated apply lost the picture fill' }
    $results.Add('Repeated apply on one Shape preserved identity and the picture fill.')

    $presentation.SaveAs($deck, 24)
    $results.Add('SaveAs succeeded.')
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

$reopened = $app.Presentations.Open($deck, $true, $false, $false)
try {
    $withPicture = 0
    $pictureShapes = 0
    foreach ($slide in $reopened.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ([int]$shape.Type -eq 13) { $pictureShapes++ }
            if ([int]$shape.Fill.Type -eq 6) { $withPicture++ }
        }
    }
    $results.Add("Reopen: $withPicture Shapes carry a picture fill; Picture shapes: $pictureShapes")
    if ($withPicture -lt 6) { throw 'Not every Shape kept its picture fill after reopen' }
    if ($pictureShapes -ne 0) { throw 'Reopened deck contains Picture shapes' }
} finally {
    $reopened.Saved = -1
    $reopened.Close()
}

$results | Set-Content "$root/artifacts/native_apply_reuse.txt"
$results
