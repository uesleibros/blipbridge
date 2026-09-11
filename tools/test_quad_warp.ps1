<#
.SYNOPSIS
Proves the projective warp end to end: does PowerPoint render the perspective?

.DESCRIPTION
Two things were established separately, and neither is worth much alone:

  - PowerPoint maps a picture fill onto the Shape's bounding box and clips it to
    the path (tools/probe_freeform_uv_mapping.ps1)
  - src/image/warp.cpp produces a correct projective mapping
    (tests/warp_contract.cpp, which never opens PowerPoint)

This does both at once and reads the result out of what Office actually drew.
The image is a UV gradient, so a rendered pixel says which source texel arrived
there. For each quad:

  1. build a Freeform with those four nodes
  2. warp the UV image onto the same four points and apply it
  3. export the Shape and read the UVs back

Then check the property that matters. Without the warp, the UV at a point inside
the shape follows the *bounding box* - measured in the earlier probe, to within
a texel. With the warp, it has to follow the *quad*: the source corners must land
on the quad's corners, and for a trapezoid the middle of the texture must sit
nearer the narrow end than half way, which is what perspective means and what an
affine map would get wrong.

A test that only checked "the corners are right" would pass for an affine
transform too, so the foreshortening check is the one that earns the feature.
#>
param([int]$Probe = 256, [switch]$KeepImages)

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

$results = New-Object System.Collections.Generic.List[string]
$failures = 0
function Assert([bool]$condition, [string]$what) {
    if ($condition) { $script:results.Add("  ok   $what") }
    else { $script:results.Add("  FAIL $what"); $script:failures++ }
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ("bb_quad_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

# R = u, G = v: the image carries its own coordinates.
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


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $originX = 60.0
    $originY = 60.0

    # A strong taper: the top edge is a quarter of the bottom, so perspective and
    # affine disagree by far more than any sampling noise. Kept as two flat
    # arrays because @() unrolls nested array literals and $quad[0][0] would then
    # not mean what it reads like.
    $qx = @(120.0, 200.0, 280.0, 0.0)
    $qy = @(0.0, 0.0, 220.0, 220.0)

    $builder = $slide.Shapes.BuildFreeform(1, $originX + $qx[0], $originY + $qy[0])
    foreach ($index in 1..3) {
        $null = $builder.AddNodes(0, 0, $originX + $qx[$index], $originY + $qy[$index])
    }
    $null = $builder.AddNodes(0, 0, $originX + $qx[0], $originY + $qy[0])
    $shape = $builder.ConvertToShape()
    $shape.Name = 'WarpedQuad'
    $shape.Line.Visible = $msoFalse

    # The quad handed to the warp is in the same coordinates as the nodes, so the
    # warped raster covers exactly the Shape's bounding box.
    [double[]]$points = @(
        ($originX + $qx[0]), ($originY + $qy[0]),
        ($originX + $qx[1]), ($originY + $qy[1]),
        ($originX + $qx[2]), ($originY + $qy[2]),
        ($originX + $qx[3]), ($originY + $qy[3]))

    $report = $engine.WarpApplyQuad($shape, $uvPath, $points)
    $results.Add("  note $report")
    Assert ($shape.Fill.Type -eq 6) 'the warped image is applied as a picture fill'

    $exported = Join-Path $temp 'warped.png'
    $shape.Export($exported, $ppShapeFormatPNG)
    $render = [System.Drawing.Bitmap]::FromFile($exported)
    try {
        $width = $render.Width
        $height = $render.Height
        $results.Add("  note render $width x $height")

        function UVAt([int]$x, [int]$y) {
            $pixel = $render.GetPixel($x, $y)
            if ($pixel.A -lt 250) { return $null }
            return @{ u = $pixel.R / 255.0; v = $pixel.G / 255.0 }
        }

        # --- the corners follow the quad, not the box -----------------------
        # Sampled a little inside each node, because the node itself is on the
        # boundary where antialiasing decides.
        $topLeftNode = UVAt ([int]($width * ($qx[0] + 12) / 280.0)) ([int]($height * 0.04))
        $topRightNode = UVAt ([int]($width * ($qx[1] - 12) / 280.0)) ([int]($height * 0.04))
        Assert ($null -ne $topLeftNode -and $topLeftNode.u -lt 0.25) `
            'the source left edge arrives at the quad top-left node, not the box corner'
        Assert ($null -ne $topRightNode -and $topRightNode.u -gt 0.75) `
            'and the source right edge arrives at the quad top-right node'

        # Without the warp, the box mapping would put u = x/width at the top of
        # the box - which at the narrow top edge is roughly 0.43 to 0.71, not 0
        # and 1. So the two checks above are exactly what distinguishes them.

        # --- the foreshortening is projective, not affine -------------------
        # Walk down the centre line and find where the source mid-row lands.
        $midRow = -1
        for ($y = 0; $y -lt $height; $y++) {
            $sample = UVAt ([int]($width / 2)) $y
            if ($null -ne $sample -and $sample.v -ge 0.5) { $midRow = $y; break }
        }
        $fraction = if ($midRow -ge 0) { $midRow / [double]$height } else { [double]::NaN }
        $results.Add("  note the source mid-row lands at {0:P0} of the height" -f $fraction)
        Assert ($midRow -ge 0) 'the source mid-row is somewhere on the centre line'
        Assert ($midRow -ge 0 -and $fraction -lt 0.42) `
            'and it sits well above half height, so the render really is perspective (affine would be 50%)'

        # --- and it is monotonic, with no folding ---------------------------
        $monotonic = $true
        $previous = -1.0
        for ($y = 0; $y -lt $height; $y++) {
            $sample = UVAt ([int]($width / 2)) $y
            if ($null -eq $sample) { continue }
            if ($sample.v -lt $previous - 0.02) { $monotonic = $false; break }
            $previous = $sample.v
        }
        Assert $monotonic 'v increases monotonically down the quad, with no folding or wrap'

        # --- outside the quad stays empty -----------------------------------
        $outside = $render.GetPixel(2, 2)
        Assert ($outside.A -lt 250) 'the bounding-box corner outside the quad is not painted'
    } finally {
        $render.Dispose()
    }

    # --- it survives save and reopen ---------------------------------------
    $saved = Join-Path $temp 'quad.pptx'
    $presentation.SaveAs($saved, 11)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopened = $presentation.Slides.Item(1).Shapes.Item('WarpedQuad')
    Assert ($reopened.Fill.Type -eq 6) 'the warped fill survives save and reopen'
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
    if (-not $KeepImages) { Remove-Item $temp -Recurse -Force -ErrorAction SilentlyContinue }
}

$results
if ($KeepImages) { "renders kept in $temp" }
if ($failures -gt 0) { throw "$failures check(s) failed" }
"all $($results.Count) lines, no failures"
