<#
.SYNOPSIS
First controlled native picture-fill apply, without Fill.UserPicture.

.DESCRIPTION
Applies image bytes to an existing normal AutoShape through the native path:
memory IStream -> GEL::ICachedImage::Create -> OART record -> transaction ->
receiver. No donor Shape, no PickUp/Apply, no UserPicture on the target and no
source file for the applied image.

The script validates Shape identity, geometry, Z order, type and fill, then
saves, closes, reopens and re-checks. A separate reference Shape filled the
ordinary way gives a pixel comparison.

GFX and OART are warmed with one ordinary picture fill on a throwaway Shape that
is deleted before the target is touched; the target Shape itself never sees
UserPicture.
#>
$ErrorActionPreference = 'Stop'

<#
Connecting can land on a PowerPoint that a previous suite is still shutting
down, which fails with 0x800706B5 "unknown interface". That says nothing about
BlipBridge, so the connection waits for the dying host and retries rather than
reporting a failure the code did not cause.
#>
function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        try {
            $candidate = New-Object -ComObject PowerPoint.Application
            # Activation succeeding is not the same as the host being usable. A
            # PowerPoint that is part-way through quitting will hand back an
            # object whose properties then fail, and the suite dies later with
            # something that looks like a BlipBridge bug - "the object did not
            # answer Shape.Type" - rather than like the teardown race it is. So
            # ask it something before trusting it.
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
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]
$deck = Join-Path $root 'artifacts/native_apply.pptx'

function Get-Identity($shape) {
    @($shape.Id, $shape.Name, $shape.Type, $shape.Left, $shape.Top, $shape.Width,
      $shape.Height, $shape.Rotation, $shape.ZOrderPosition) -join '|'
}


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, 12)

    # Warm GFX/OART on a throwaway Shape, then remove it. The target never sees
    # UserPicture at any point.
    $warmup = $slide.Shapes.AddShape(1, 300, 200, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    $target = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
    $target.Name = 'BB_native_target'
    $before = Get-Identity $target
    $fillTypeBefore = [int]$target.Fill.Type

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    $report = $engine.NativeApplyExperiment($target.Fill, $bytes)
    $results.Add("apply report: $report")

    $after = Get-Identity $target
    if ($after -ne $before) { throw "Shape identity or geometry changed:`n  $before`n  $after" }
    $results.Add('Shape identity, name, geometry, rotation and Z order unchanged.')

    $fillType = [int]$target.Fill.Type
    $results.Add("Fill.Type before=$fillTypeBefore after=$fillType (6 is picture)")
    if ($fillType -ne 6) { throw "Native apply did not produce a picture fill (Fill.Type=$fillType)" }

    if ([int]$target.Type -ne 1) { throw 'Shape is no longer an AutoShape' }
    $results.Add('Shape is still a normal AutoShape, not a Picture shape.')

    # Reference Shape filled the ordinary way, for a pixel comparison.
    $reference = $slide.Shapes.AddShape(1, 200, 20, 120, 120)
    $reference.Fill.UserPicture($texture)
    $target.Export("$root/artifacts/native_apply_shape.png", 2)
    $reference.Export("$root/artifacts/native_apply_reference.png", 2)
    $reference.Delete()

    # Freeform variant: nodes must survive the apply and stay editable.
    $builder = $slide.Shapes.BuildFreeform(0, 400, 300)
    $builder.AddNodes(0, 0, 540, 300)
    $builder.AddNodes(0, 0, 470, 440)
    $builder.AddNodes(0, 0, 400, 300)
    $freeform = $builder.ConvertToShape()
    $freeform.Name = 'BB_native_freeform'
    $freeformBefore = Get-Identity $freeform
    $nodesBefore = $freeform.Nodes.Count
    $null = $engine.NativeApplyExperiment($freeform.Fill, $bytes)
    if ((Get-Identity $freeform) -ne $freeformBefore) { throw 'Freeform identity changed' }
    if ($freeform.Nodes.Count -ne $nodesBefore) { throw 'Freeform node count changed' }
    $freeform.Nodes.SetPosition(2, 560, 300)
    if ($freeform.Nodes.Count -ne $nodesBefore) { throw 'Freeform nodes stopped being editable' }
    if ([int]$freeform.Fill.Type -ne 6) { throw 'Freeform did not get a picture fill' }
    $results.Add("Freeform kept $nodesBefore editable nodes and got a picture fill.")

    $presentation.SaveAs($deck, 24)
    $results.Add('SaveAs succeeded.')
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

# Reopen and re-check persistence.
$reopened = $app.Presentations.Open($deck, $true, $false, $false)
try {
    $slide = $reopened.Slides.Item(1)
    $picture = 0
    foreach ($shape in $slide.Shapes) {
        if ([int]$shape.Type -eq 13) { $picture++ }
    }
    $reloaded = $slide.Shapes | Where-Object { $_.Name -eq 'BB_native_target' } | Select-Object -First 1
    if (-not $reloaded) { throw 'Target Shape missing after reopen' }
    if ([int]$reloaded.Fill.Type -ne 6) { throw 'Picture fill did not survive reopen' }
    if ([int]$reloaded.Type -ne 1) { throw 'Reopened Shape is not an AutoShape' }
    $reloaded.Export("$root/artifacts/native_apply_reopened.png", 2)
    $results.Add("Reopen preserved the picture fill; Picture shapes on slide: $picture")
} finally {
    $reopened.Saved = -1
    $reopened.Close()
}

Add-Type -AssemblyName System.Drawing
function Compare-Png($left, $right) {
    $a = [System.Drawing.Bitmap]::FromFile($left)
    $b = [System.Drawing.Bitmap]::FromFile($right)
    try {
        if ($a.Width -ne $b.Width -or $a.Height -ne $b.Height) { return -1 }
        $differences = 0
        for ($y = 0; $y -lt $a.Height; $y += 2) {
            for ($x = 0; $x -lt $a.Width; $x += 2) {
                if ($a.GetPixel($x, $y).ToArgb() -ne $b.GetPixel($x, $y).ToArgb()) { $differences++ }
            }
        }
        return $differences
    } finally { $a.Dispose(); $b.Dispose() }
}

$againstReference = Compare-Png "$root/artifacts/native_apply_shape.png" "$root/artifacts/native_apply_reference.png"
$againstReopened = Compare-Png "$root/artifacts/native_apply_shape.png" "$root/artifacts/native_apply_reopened.png"
$results.Add("Sampled pixel differences against an ordinary UserPicture fill: $againstReference")
$results.Add("Sampled pixel differences against the reopened Shape: $againstReopened")
if ($againstReference -ne 0) { $results.Add('WARNING: native fill does not match the ordinary fill') }
if ($againstReopened -ne 0) { $results.Add('WARNING: reopened fill does not match') }

$results | Set-Content "$root/artifacts/native_apply.txt"
$results
