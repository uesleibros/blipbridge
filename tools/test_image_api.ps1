<#
.SYNOPSIS
Verifies the ABI 5 image surface in PowerPoint, through the real C ABI.

.DESCRIPTION
The design this asserts is that there are two resources and they stay separate:

  BlipBridgeImage    decoded BGRA that BlipBridge owns, on the CPU, so it can be
                     cropped, oriented, scaled and warped repeatedly
  BlipBridgeTexture  what Office holds, with no pixels anyone can reach

So the checks are about the seam as much as the pixels: an image handle is not a
texture handle and each refuses the other; LoadTextureScaled never creates an
image; ApplyImageQuad owns the texture it makes and leaves none behind; and image
processing on its own never touches the document, so it creates no undo entry -
only applying to a Shape does.

Every call goes through the shipped exports via a thin research forwarder, not a
re-implementation.
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
$msoTrue = -1
$msoFalse = 0
$ppLayoutBlank = 12
$ppSaveAsDefault = 11

Add-Type -AssemblyName System.Drawing

$results = New-Object System.Collections.Generic.List[string]
$failures = 0
function Assert([bool]$condition, [string]$what) {
    if ($condition) { $script:results.Add("  ok   $what") }
    else { $script:results.Add("  FAIL $what"); $script:failures++ }
}
function Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ("bb_img_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

# A 64x64 atlas of four 32x32 quadrants, so a crop can be checked by its colour.
$atlasPath = Join-Path $temp 'atlas.png'
$atlas = New-Object System.Drawing.Bitmap(64, 64, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
for ($y = 0; $y -lt 64; $y++) {
    for ($x = 0; $x -lt 64; $x++) {
        $colour = if ($x -lt 32 -and $y -lt 32) { [System.Drawing.Color]::FromArgb(255, 255, 0, 0) }
                  elseif ($y -lt 32) { [System.Drawing.Color]::FromArgb(255, 0, 255, 0) }
                  elseif ($x -lt 32) { [System.Drawing.Color]::FromArgb(255, 0, 0, 255) }
                  else { [System.Drawing.Color]::FromArgb(255, 255, 255, 0) }
        $atlas.SetPixel($x, $y, $colour)
    }
}
$atlas.Save($atlasPath, [System.Drawing.Imaging.ImageFormat]::Png)
$atlas.Dispose()

$saved = Join-Path $temp 'image_api.pptx'

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    # --- the capability advertises the surface -------------------------------
    $mask = [uint32](Field ($engine.AbiCapabilities()) 'mask')
    Assert (($mask -band 0x200) -ne 0) `
        ("BB_CAP_IMAGE_PIPELINE is set in host (mask 0x{0:X4})" -f $mask)

    # --- loading an image ----------------------------------------------------
    $loaded = $engine.LoadImageAbi($atlasPath, '')
    $image = [long](Field $loaded 'image')
    Assert ((Field $loaded 'width') -eq '64' -and (Field $loaded 'height') -eq '64') `
        'an image loads at its own size'
    # The first handle *is* the base, so this is -ge. Texture handles start at
    # 0x1000000 and the spaces are far apart precisely so this can be checked.
    Assert ($image -ge 0x4000000) 'and its handle comes from the image space, not the texture space'

    # Counts are deltas, never absolutes. The store lives as long as the
    # PowerPoint process, so a second run of this harness starts with whatever
    # the first one left - and an absolute count would fail for a reason that has
    # nothing to do with the code.
    $baseline = [int](Field $loaded 'count')
    Assert ($baseline -ge 1) 'and the image count includes it'

    # --- crop, orient, resize on the way in ----------------------------------
    # cropX,cropY,cropW,cropH,transform,targetW,targetH,filter
    $tile = $engine.LoadImageAbi($atlasPath, '32,0,32,32,0,128,128,0')
    Assert ((Field $tile 'width') -eq '128' -and (Field $tile 'height') -eq '128') `
        'a 32x32 region scales to 128x128 on the way in'

    $rotated = $engine.LoadImageAbi($atlasPath, '0,0,64,32,3,0,0,0')
    Assert ((Field $rotated 'width') -eq '32' -and (Field $rotated 'height') -eq '64') `
        'rotating a 64x32 crop by 90 transposes it to 32x64'

    $bad = $false
    try { $null = $engine.LoadImageAbi($atlasPath, '0,0,999,999,0,0,0,0') } catch { $bad = $true }
    Assert $bad 'a region reaching outside the image is refused'

    $badFilter = $false
    try { $null = $engine.LoadImageAbi($atlasPath, '0,0,0,0,0,32,32,99') } catch { $badFilter = $true }
    Assert $badFilter 'an unknown filter is refused rather than clamped'

    $missing = $false
    try { $null = $engine.LoadImageAbi((Join-Path $temp 'nope.png'), '') } catch { $missing = $true }
    Assert $missing 'a missing file is refused'

    # --- the two resources stay separate -------------------------------------
    $scaled = $engine.LoadTextureScaledAbi($atlasPath, '0,0,0,0,0,16,16,0')
    $textureHandle = [long](Field $scaled 'texture')
    Assert ($textureHandle -gt 0) 'LoadTextureScaled produces a texture'
    Assert ([int](Field $scaled 'images') -eq $baseline + 2) `
        'and creates no CPU image - only the two loaded since the baseline were added'

    $wrongKind = $false
    try { $null = $engine.ReleaseImageAbi($textureHandle) } catch { $wrongKind = $true }
    Assert $wrongKind 'a texture handle passed to ReleaseImage is refused, not mistaken for an image'

    # --- the quad path -------------------------------------------------------
    $originX = 60.0
    $originY = 60.0
    $qx = @(120.0, 200.0, 280.0, 0.0)
    $qy = @(0.0, 0.0, 220.0, 220.0)
    $builder = $slide.Shapes.BuildFreeform(1, $originX + $qx[0], $originY + $qy[0])
    foreach ($index in 1..3) {
        $null = $builder.AddNodes(0, 0, $originX + $qx[$index], $originY + $qy[$index])
    }
    $null = $builder.AddNodes(0, 0, $originX + $qx[0], $originY + $qy[0])
    $quadShape = $builder.ConvertToShape()
    $quadShape.Name = 'QuadShape'
    $quadShape.Line.Visible = $msoFalse

    [double[]]$points = @(
        ($originX + $qx[0]), ($originY + $qy[0]), ($originX + $qx[1]), ($originY + $qy[1]),
        ($originX + $qx[2]), ($originY + $qy[2]), ($originX + $qx[3]), ($originY + $qy[3]))

    $texturesBefore = $engine.GetTextureCount()
    $applied = $engine.ApplyImageQuadAbi($quadShape, $image, $points)
    Assert ($quadShape.Fill.Type -eq 6) 'ApplyImageQuad leaves a picture fill'
    Assert ([int](Field $applied 'textures') -eq $texturesBefore) `
        'and releases the texture it made, so the count is unchanged'

    # Applying the same image again must work: the image is not consumed.
    $null = $engine.ApplyImageQuadAbi($quadShape, $image, $points)
    Assert ($quadShape.Fill.Type -eq 6) 'the image survives being warped and applied twice'

    # --- undo: only the apply touches the document ---------------------------
    # A marker Shape goes on the undo stack, then images are loaded and released.
    # If any of that reached the document, the marker would need more than one
    # Undo to remove.
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()
    $marker = $slide.Shapes.AddShape(1, 400, 400, 40, 40)
    $marker.Name = 'ImageMarker'
    $app.StartNewUndoEntry()

    $throwaway = $engine.LoadImageAbi($atlasPath, '0,0,32,32,1,64,64,1')
    $null = $engine.ReleaseImageAbi([long](Field $throwaway 'image'))

    $app.CommandBars.ExecuteMso('Undo')
    $markerGone = $true
    try { $null = $slide.Shapes.Item('ImageMarker'); $markerGone = $false } catch { }
    Assert $markerGone 'loading and releasing images creates no undo entry at all'

    # --- release semantics ---------------------------------------------------
    $released = $engine.ReleaseImageAbi($image)
    Assert ($null -ne $released) 'an image releases'
    $stale = $false
    try { $null = $engine.ReleaseImageAbi($image) } catch { $stale = $true }
    Assert $stale 'and releasing it twice is refused - handles are never recycled'

    # --- persistence ---------------------------------------------------------
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopened = $presentation.Slides.Item(1).Shapes.Item('QuadShape')
    Assert ($reopened.Fill.Type -eq 6) 'a quad-warped fill survives save and reopen'
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
    Remove-Item $temp -Recurse -Force -ErrorAction SilentlyContinue
}

$results
if ($failures -gt 0) { throw "$failures check(s) failed" }
"all $($results.Count) checks passed"
