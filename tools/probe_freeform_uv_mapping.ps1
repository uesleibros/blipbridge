<#
.SYNOPSIS
Finds out how PowerPoint maps a picture fill onto a four-node Freeform.

.DESCRIPTION
Everything about a quad/perspective API depends on one question that has to be
answered by measurement rather than assumed: when a picture fills a Freeform,
what decides where each part of the image lands?

Two hypotheses, and they predict different pixels:

  A. bounding box.  The image is mapped linearly onto the Shape's bounding
     rectangle and then clipped to the path. The geometry of the path changes
     what is *visible*, not where the texture sits. Prediction: at a point (x, y)
     inside the exported image, u = x / width and v = y / height, whatever shape
     the path is.

  B. quad corners.  The image's four corners are mapped to the path's four
     nodes, so the texture follows the edges. Prediction: u and v come from the
     inverse of that mapping, and for anything other than a rectangle they are
     nowhere near x / width.

The probe texture is a UV gradient - red carries u, green carries v - so a single
exported pixel says which part of the source image arrived there. No inference
from colours "looking right": the image encodes its own coordinates.

Each Freeform is exported to PNG, sampled at a grid of interior points, and each
sample is scored against hypothesis A. The report is the mean and worst error in
source-pixel units, per shape.

.NOTES
This decides whether pre-warping can work at all. If A holds, a caller-supplied
quad can be compensated for, because we would know exactly what Office will do
with the image. If neither holds, a quad API is not something to design yet.
#>
param([int]$Probe = 512, [switch]$KeepImages)

$ErrorActionPreference = 'Stop'

<#
Connecting can land on a PowerPoint that a previous suite is still shutting
down, which fails with 0x800706B5 "unknown interface". That says nothing about
BlipBridge, so the connection waits for the dying host and retries rather than
reporting a failure the code did not cause.
#>
function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try { return New-Object -ComObject PowerPoint.Application }
        catch {
            if ($attempt -eq 10) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$msoFalse = 0
$ppLayoutBlank = 12
$ppShapeFormatPNG = 2

Add-Type -AssemblyName System.Drawing

$temp = Join-Path ([IO.Path]::GetTempPath()) ("bb_uv_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

# --- the probe texture ------------------------------------------------------
# R = u * 255, G = v * 255, B = 128. A pixel read back from the render says
# exactly which source texel it came from, to within a quantisation step.
$uvPath = Join-Path $temp 'uv.png'
$bitmap = New-Object System.Drawing.Bitmap($Probe, $Probe, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
for ($y = 0; $y -lt $Probe; $y++) {
    for ($x = 0; $x -lt $Probe; $x++) {
        $r = [int](255.0 * $x / ($Probe - 1))
        $g = [int](255.0 * $y / ($Probe - 1))
        $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, $r, $g, 128))
    }
}
$bitmap.Save($uvPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()


<#
The quads under test. Each is a list of four points in the order the API will
document - source top-left, top-right, bottom-right, bottom-left - given in
points relative to the shape's origin.
#>
$cases = [ordered]@{
    'rectangle'        = @(@(0, 0), @(200, 0), @(200, 200), @(0, 200))
    'wide rectangle'   = @(@(0, 0), @(300, 0), @(300, 120), @(0, 120))
    'parallelogram'    = @(@(60, 0), @(260, 0), @(200, 200), @(0, 200))
    'trapezoid'        = @(@(50, 0), @(150, 0), @(200, 200), @(0, 200))
    'extreme trapezoid'= @(@(90, 0), @(110, 0), @(200, 200), @(0, 200))
    'rotated quad'     = @(@(100, 0), @(200, 100), @(100, 200), @(0, 100))
    'perspective quad' = @(@(40, 20), @(180, 0), @(200, 200), @(0, 160))
}

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$rows = @()
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    [byte[]]$bytes = [IO.File]::ReadAllBytes($uvPath)
    $handle = $engine.LoadTexture($bytes)

    foreach ($name in $cases.Keys) {
        $points = $cases[$name]
        $originX = 40.0
        $originY = 40.0

        $builder = $slide.Shapes.BuildFreeform(1, $originX + $points[0][0], $originY + $points[0][1])
        foreach ($index in 1..3) {
            $null = $builder.AddNodes(0, 0, $originX + $points[$index][0], $originY + $points[$index][1])
        }
        $null = $builder.AddNodes(0, 0, $originX + $points[0][0], $originY + $points[0][1])
        $shape = $builder.ConvertToShape()
        $shape.Name = "Quad_$($name -replace '\s','_')"
        $shape.Line.Visible = $msoFalse

        $null = $engine.ApplyTexture($shape, $handle)

        $exported = Join-Path $temp ("$($shape.Name).png")
        $shape.Export($exported, $ppShapeFormatPNG)

        # --- score the render against hypothesis A -------------------------
        $render = [System.Drawing.Bitmap]::FromFile($exported)
        try {
            $width = $render.Width
            $height = $render.Height
            $errors = New-Object System.Collections.Generic.List[double]
            $opaque = 0

            for ($sy = 1; $sy -lt 10; $sy++) {
                for ($sx = 1; $sx -lt 10; $sx++) {
                    $px = [int]($width * $sx / 10.0)
                    $py = [int]($height * $sy / 10.0)
                    if ($px -ge $width -or $py -ge $height) { continue }
                    $pixel = $render.GetPixel($px, $py)
                    # Only sample where the fill actually painted. Outside the
                    # path the export is transparent, and a cleared pixel says
                    # nothing about the mapping.
                    if ($pixel.A -lt 250) { continue }
                    $opaque++

                    $u = $pixel.R / 255.0
                    $v = $pixel.G / 255.0
                    $predictedU = $px / [double]($width - 1)
                    $predictedV = $py / [double]($height - 1)
                    # Express the miss in source pixels, which is the unit a
                    # reader can judge: 2 is invisible, 40 is a different image.
                    $errors.Add([math]::Sqrt(
                        [math]::Pow(($u - $predictedU) * $Probe, 2) +
                        [math]::Pow(($v - $predictedV) * $Probe, 2)))
                }
            }

            $mean = if ($errors.Count -gt 0) { ($errors | Measure-Object -Average).Average } else { [double]::NaN }
            $worst = if ($errors.Count -gt 0) { ($errors | Measure-Object -Maximum).Maximum } else { [double]::NaN }
            $rows += [pscustomobject]@{
                Quad          = $name
                RenderPx      = "$width x $height"
                Samples       = $opaque
                MeanErrorPx   = [math]::Round($mean, 1)
                WorstErrorPx  = [math]::Round($worst, 1)
                MatchesBoxMap = if ($worst -lt 12) { 'yes' } else { 'no' }
            }
        } finally {
            $render.Dispose()
        }
        $shape.Delete()
    }

    $null = $engine.ReleaseTexture($handle)
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

"probe texture: ${Probe}x${Probe} UV gradient; errors are in source-texel units"
$rows | Format-Table -AutoSize
@'
Hypothesis A - the image is mapped linearly onto the Shape's bounding box and
then clipped to the path - predicts a near-zero error for every quad, including
the ones that are nothing like a rectangle. A large error on the non-rectangular
quads would mean the path geometry itself moves the texture, which is
hypothesis B.
'@

if ($KeepImages) { "renders kept in $temp" } else { Remove-Item $temp -Recurse -Force }
