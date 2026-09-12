<#
.SYNOPSIS
The whole public BlipBridge surface, exercised in a real PowerPoint, against
whichever backend is loaded.

.DESCRIPTION
BlipBridge has two Office backends: the accelerated one that drives Office
internals on x64, and the portable one that drives documented Automation. The
promise of v0.8 is that a caller uses the same API either way, and the only
honest way to hold that promise is to run the same suite against both.

So this suite is backend-aware rather than backend-specific. It reads which
backend is loaded, asserts the capability mask that backend must report, and then
asserts *identical behaviour* for everything else. Run it twice - once against a
normal x64 build and once against one configured with
BB_FORCE_PORTABLE_BACKEND=ON - and a difference in the results is a difference in
the API, which is exactly what must not exist.

## What this does and does not prove

It proves the portable backend's behaviour inside a real PowerPoint: real
Application, Presentation, Slide, Shape, ShapeRange, Fill, Undo, and SaveAs and
reopen. Every Shape is real and every fill is a real document edit.

It does **not** prove anything about 32-bit PowerPoint. The portable backend is
the backend an x86 build uses, and the only thing x86-specific about it is the
compiler - but that is a reason to expect it to work there, not evidence that it
does. See docs/windows_x86.md.

## Assertions

Every check asserts the specific thing its name claims. Where a result is
visible, it is checked by rendering the Shape and sampling pixels rather than by
asking a cache what it thinks it did - a cache that is wrong will answer
confidently.

.PARAMETER KeepArtifacts
Leaves the rendered PNGs and the saved presentation behind for inspection.
#>
param([switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'

<#
Connecting can land on a PowerPoint that a previous suite is still shutting down,
which fails with 0x800706B5. That says nothing about BlipBridge, so the
connection waits for the dying host and retries. Activation succeeding is also
not the same as the host being usable, so it is asked something before it is
trusted.
#>
function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        try {
            $candidate = New-Object -ComObject PowerPoint.Application
            $null = $candidate.Version
            $null = $candidate.Presentations.Count
            return $candidate
        } catch {
            if ($attempt -eq 15) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

Add-Type -AssemblyName System.Drawing

$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$ppSaveAsDefault = 11
$ppShapeFormatPNG = 2
$msoShapeRectangle = 1
$msoFillPicture = 6

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
function Refused([scriptblock]$action) {
    try { & $action | Out-Null; return '' } catch { return ($_.Exception.Message -replace '\s+', ' ') }
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('bb_portable_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

<#
Deterministic single-colour images, so "which picture is on this Shape" is a
question a sampled pixel can answer. A photograph would make the same assertion
depend on scaling and JPEG noise.
#>
function New-SolidPng([string]$path, [int]$r, [int]$g, [int]$b, [int]$size = 64) {
    $bitmap = New-Object System.Drawing.Bitmap($size, $size,
        [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, $r, $g, $b))
    $graphics.FillRectangle($brush, 0, 0, $size, $size)
    $graphics.Dispose(); $brush.Dispose()
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

# Straight BGRA32, top-down, tightly packed: the canonical format the raw-pixel
# entry points document.
function New-BgraBytes([int]$width, [int]$height, [int]$r, [int]$g, [int]$b) {
    $bytes = New-Object 'byte[]' ($width * $height * 4)
    for ($i = 0; $i -lt ($width * $height); $i++) {
        $bytes[$i * 4] = [byte]$b
        $bytes[$i * 4 + 1] = [byte]$g
        $bytes[$i * 4 + 2] = [byte]$r
        $bytes[$i * 4 + 3] = 255
    }
    return , $bytes
}

<#
Renders a Shape and reports the colour at its centre. This is the only way to
answer "which image is actually on it": Fill.Type says a picture is there and
never which one, and a stale cache will claim the right answer while showing the
wrong picture. That failure is the whole reason several checks below exist.
#>
function Get-ShapeColour($shape, [string]$tag) {
    $png = Join-Path $temp "$tag.png"
    $shape.Export($png, $ppShapeFormatPNG)
    $bitmap = [System.Drawing.Bitmap]::FromFile($png)
    try {
        $pixel = $bitmap.GetPixel([int]($bitmap.Width / 2), [int]($bitmap.Height / 2))
        return [pscustomobject]@{ R = $pixel.R; G = $pixel.G; B = $pixel.B; Path = $png }
    } finally { $bitmap.Dispose() }
}
function Near([int]$actual, [int]$expected, [int]$tolerance = 40) {
    return [Math]::Abs($actual - $expected) -le $tolerance
}
function IsRed($c) { return (Near $c.R 220 60) -and (Near $c.G 30 60) -and (Near $c.B 30 60) }
function IsBlue($c) { return (Near $c.R 30 60) -and (Near $c.G 30 60) -and (Near $c.B 220 60) }
function IsGreen($c) { return (Near $c.R 30 60) -and (Near $c.G 200 70) -and (Near $c.B 30 60) }

$redPath = Join-Path $temp 'red.png'
$bluePath = Join-Path $temp 'blue.png'
$rewritePath = Join-Path $temp 'rewritten.png'
New-SolidPng $redPath 220 30 30
New-SolidPng $bluePath 30 30 220
New-SolidPng $rewritePath 220 30 30       # starts as red; becomes blue later

$saved = Join-Path $temp 'portable_matrix.pptx'

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    # --- 1-4: initialize, availability, backend, capabilities -----------------
    # Every call below implies Initialize; asking for the capability mask is the
    # first thing that would fail if it had not happened.
    $report = $engine.AbiCapabilities()
    $mask = [uint32](Field $report 'mask')
    $backend = Field $report 'backend'
    $accelerated = $backend -eq 'windows-office-native'
    $results.Add("  note backend: $backend  mask: 0x$('{0:X4}' -f $mask)")

    Assert ($mask -ne 0) 'Initialize succeeded and the library reports capabilities'

    # IsAvailable is "can BlipBridge do its job here", which is a non-zero mask.
    # It must not be a test of the accelerated bit - that is IsAccelerated.
    Assert ($mask -ne 0) 'IsAvailable: the host is one BlipBridge can serve'
    Assert ((($mask -band 0x1) -ne 0) -eq $accelerated) `
        "IsAccelerated agrees with the loaded backend ($backend)"

    if ($accelerated) {
        Assert ($mask -eq 0x03FF) ('the accelerated mask is 0x03FF (got 0x{0:X4})' -f $mask)
    } else {
        Assert ($mask -eq 0x03FE) ('the portable mask is 0x03FE (got 0x{0:X4})' -f $mask)
        Assert (($mask -band 0x1) -eq 0) 'and the only bit it drops is BB_CAP_NATIVE_BACKEND'
    }

    # --- 5: LoadTexture from encoded bytes ------------------------------------
    [byte[]]$redBytes = [IO.File]::ReadAllBytes($redPath)
    [byte[]]$blueBytes = [IO.File]::ReadAllBytes($bluePath)
    $loaded = $engine.LoadTextureBytesAbi($redBytes)
    $redTexture = [long](Field $loaded 'texture')
    Assert ($redTexture -ge 0x1000000) 'LoadTexture returns a handle from the texture space'

    # --- 6: LoadTexturePixels from raw BGRA -----------------------------------
    $pixels = New-BgraBytes 32 32 30 200 30
    $loaded = $engine.LoadTexturePixelsAbi($pixels, 32, 32, 128)
    $greenTexture = [long](Field $loaded 'texture')
    Assert ($greenTexture -ge 0x1000000 -and $greenTexture -ne $redTexture) `
        'LoadTexturePixels returns a distinct texture handle'

    # --- 7-8: LoadTextureScaled, and from a file ------------------------------
    # cropX,cropY,cropW,cropH,transform,targetW,targetH,filter
    $scaled = $engine.LoadTextureScaledAbi($redPath, '0,0,0,0,0,128,128,1')
    $scaledTexture = [long](Field $scaled 'texture')
    Assert ($scaledTexture -ge 0x1000000) 'LoadTextureScaledFromFile returns a texture'
    Assert ((Field $scaled 'images') -eq (Field (($engine.AbiCounts())) 'images')) `
        'and it creates no CPU image, which is why that path exists'

    # --- 9-12: images ---------------------------------------------------------
    $imageReport = $engine.LoadImageAbi($redPath, '')
    $redImage = [long](Field $imageReport 'image')
    Assert ($redImage -ge 0x4000000) 'LoadImageFromFile returns a handle from the image space'
    Assert ((Field $imageReport 'width') -eq '64' -and (Field $imageReport 'height') -eq '64') `
        'GetImageSize reports the image it loaded'

    $pixelImage = $engine.LoadImagePixelsAbi((New-BgraBytes 16 8 30 30 220), 16, 8, 64)
    $bluePixelImage = [long](Field $pixelImage 'image')
    Assert ((Field $pixelImage 'width') -eq '16' -and (Field $pixelImage 'height') -eq '8') `
        'LoadImagePixels accepts raw BGRA and reports its size'

    $sizeReport = $engine.ImageSizeAbi($redImage)
    Assert ((Field $sizeReport 'width') -eq '64') 'GetImageSize answers for a live image handle'

    # --- 13: CreateTextureFromImage -------------------------------------------
    $fromImage = $engine.CreateTextureFromImageAbi($redImage)
    $imageTexture = [long](Field $fromImage 'texture')
    Assert ($imageTexture -ge 0x1000000) 'CreateTextureFromImage turns an image into a texture'

    # --- 16: ApplyTexture, checked by what the Shape shows --------------------
    $target = $slide.Shapes.AddShape($msoShapeRectangle, 20, 20, 120, 120)
    $target.Name = 'Target'
    $engine.ApplyTexture($target, $redTexture)
    Assert ($target.Fill.Type -eq $msoFillPicture) 'ApplyTexture leaves a picture fill'
    Assert (IsRed (Get-ShapeColour $target 'apply_red')) 'and the Shape shows that image'

    $engine.ApplyTexture($target, $greenTexture)
    Assert (IsGreen (Get-ShapeColour $target 'apply_green')) `
        'a second ApplyTexture replaces it with the new image'

    # --- 17-18: ApplyTextureIfChanged, miss then hit --------------------------
    $skipShape = $slide.Shapes.AddShape($msoShapeRectangle, 160, 20, 120, 120)
    $skipShape.Name = 'Skip'
    Assert ((Field ($engine.ApplyTextureIfChanged($skipShape, $redTexture)) 'skipped') -eq '0') `
        'ApplyTextureIfChanged does real work the first time'
    Assert (IsRed (Get-ShapeColour $skipShape 'skip_first')) 'and the image is there'
    Assert ((Field ($engine.ApplyTextureIfChanged($skipShape, $redTexture)) 'skipped') -eq '1') `
        'ApplyTextureIfChanged skips the same image on the same Shape'
    Assert (IsRed (Get-ShapeColour $skipShape 'skip_second')) `
        'and a skip leaves the right image in place, not a stale one'
    Assert ((Field ($engine.ApplyTextureIfChanged($skipShape, $greenTexture)) 'skipped') -eq '0') `
        'a different image is not skipped'
    Assert (IsGreen (Get-ShapeColour $skipShape 'skip_third')) 'and it really changes the Shape'

    # A fill changed behind BlipBridge's back must not be skipped.
    $skipShape.Fill.ForeColor.RGB = 255
    $null = $skipShape.Fill.Solid()
    Assert ((Field ($engine.ApplyTextureIfChanged($skipShape, $greenTexture)) 'skipped') -eq '0') `
        'a fill replaced with a colour is re-applied rather than skipped'
    Assert (IsGreen (Get-ShapeColour $skipShape 'skip_recovered')) 'and the picture comes back'

    # --- 19: ApplyTextureBatch ------------------------------------------------
    $batchShapes = @()
    for ($i = 0; $i -lt 4; $i++) {
        $s = $slide.Shapes.AddShape($msoShapeRectangle, 20 + $i * 70, 170, 60, 60)
        $s.Name = "Batch$i"
        $batchShapes += $s
    }
    $batchReport = $engine.ApplyTextureBatchAbi($batchShapes, $redTexture)
    Assert ((Field $batchReport 'applied') -eq '4') 'ApplyTextureBatch fills every Shape it was given'
    $batchFilled = ($batchShapes | Where-Object { $_.Fill.Type -eq $msoFillPicture }).Count
    Assert ($batchFilled -eq 4) 'and all four carry a picture fill'
    Assert (IsRed (Get-ShapeColour $batchShapes[3] 'batch_last')) 'and the last one shows the image'

    # --- 20: ApplyTextureRange ------------------------------------------------
    $rangeNames = @()
    for ($i = 0; $i -lt 5; $i++) {
        $s = $slide.Shapes.AddShape($msoShapeRectangle, 20 + $i * 70, 250, 60, 60)
        $s.Name = "Range$i"
        $rangeNames += $s.Name
    }
    $range = $slide.Shapes.Range($rangeNames)
    $rangeReport = $engine.ApplyTextureRange($range, $greenTexture)
    Assert ((Field $rangeReport 'applied') -eq '5') 'ApplyTextureRange reports every member filled'
    $rangeFilled = ($rangeNames | Where-Object { $slide.Shapes.Item($_).Fill.Type -eq $msoFillPicture }).Count
    Assert ($rangeFilled -eq 5) 'and every member really has a picture fill'
    Assert (IsGreen (Get-ShapeColour ($slide.Shapes.Item($rangeNames[2])) 'range_mid')) `
        'and a member in the middle shows the right image'

    # All or nothing: one ineligible member refuses the whole range.
    $connector = $slide.Shapes.AddConnector(1, 400, 250, 500, 300)
    $connector.Name = 'RangeConnector'
    $mixed = $slide.Shapes.Range($rangeNames + $connector.Name)
    $message = Refused { $engine.ApplyTextureRange($mixed, $redTexture) }
    Assert ($message -ne '') 'a range containing a Connector is refused'
    Assert ($message -match 'Range member') 'and the refusal names which member'
    Assert (IsGreen (Get-ShapeColour ($slide.Shapes.Item($rangeNames[0])) 'range_untouched')) `
        'and the eligible members were left exactly as they were'
    $connector.Delete()

    # --- 14-15: quad warp, and applying one ------------------------------------
    $quadShape = $slide.Shapes.AddShape($msoShapeRectangle, 320, 20, 140, 140)
    $quadShape.Name = 'Quad'
    $quad = @(0.0, 0.0, 140.0, 0.0, 110.0, 140.0, 30.0, 140.0)

    $warped = $engine.WarpImageQuadAbi($redImage, $quad, 1)
    $warpedTexture = [long](Field $warped 'texture')
    Assert ($warpedTexture -ge 0x1000000) 'WarpImageQuad produces a texture from an image'
    $engine.ApplyTexture($quadShape, $warpedTexture)
    Assert ($quadShape.Fill.Type -eq $msoFillPicture) 'and that texture applies like any other'

    $beforeQuad = [int](Field ($engine.AbiCounts()) 'textures')
    $engine.ApplyImageQuadAbi($quadShape, $redImage, $quad) | Out-Null
    $afterQuad = [int](Field ($engine.AbiCounts()) 'textures')
    Assert ($quadShape.Fill.Type -eq $msoFillPicture) 'ApplyImageQuad fills the Shape'
    Assert ($afterQuad -eq $beforeQuad) 'and releases the texture it made, leaving none behind'

    # Bicubic is refused for warping by name rather than quietly downgraded.
    $bicubic = Refused { $engine.WarpImageQuadAbi($redImage, $quad, 2) }
    Assert ($bicubic -ne '') 'WarpImageQuad refuses bicubic'
    Assert ($bicubic -match 'nearest and bilinear') 'and names the filters it does support'

    # --- 21: ApplyPicture ------------------------------------------------------
    $pictureShape = $slide.Shapes.AddShape($msoShapeRectangle, 320, 170, 120, 120)
    $pictureShape.Name = 'Picture'
    $engine.ApplyPicture($pictureShape, $bluePath) | Out-Null
    Assert ($pictureShape.Fill.Type -eq $msoFillPicture) 'ApplyPicture fills from a path'
    Assert (IsBlue (Get-ShapeColour $pictureShape 'picture_blue')) 'and the Shape shows that file'

    # --- REGRESSION 2: the same path, rewritten -------------------------------
    <#
    The failure this prevents: a caller writes a new image over the same
    temporary file, and the cache - keyed on the path alone - reports the Shape
    already carries it. The old picture stays, silently. This asserts the pixels,
    not the cache, because a cache that is wrong will answer confidently.
    #>
    $rewriteShape = $slide.Shapes.AddShape($msoShapeRectangle, 460, 170, 120, 120)
    $rewriteShape.Name = 'Rewritten'
    $engine.ApplyPicture($rewriteShape, $rewritePath) | Out-Null
    Assert (IsRed (Get-ShapeColour $rewriteShape 'rewrite_before')) `
        'a file applied by path shows its contents'

    Start-Sleep -Milliseconds 1100    # so the last-write time is visibly different
    New-SolidPng $rewritePath 30 30 220
    $engine.ApplyPicture($rewriteShape, $rewritePath) | Out-Null
    Assert (IsBlue (Get-ShapeColour $rewriteShape 'rewrite_after')) `
        'REGRESSION: rewriting the same path shows the new image, not the cached old one'

    # --- REGRESSION 1: ClearTextures then apply again -------------------------
    <#
    The portable backend writes its textures to a temporary directory, and
    clearing every texture removes it. Nothing recreated it, so every later apply
    failed with "the system cannot find the path specified". A cleared library has
    to keep working without being restarted.
    #>
    $engine.ClearTextures()
    Assert ($engine.GetTextureCount() -eq 0) 'ClearTextures drops every handle'

    $afterClear = $engine.LoadTextureBytesAbi($blueBytes)
    $afterClearTexture = [long](Field $afterClear 'texture')
    $clearShape = $slide.Shapes.AddShape($msoShapeRectangle, 460, 20, 120, 120)
    $clearShape.Name = 'AfterClear'
    $engine.ApplyTexture($clearShape, $afterClearTexture)
    Assert ($clearShape.Fill.Type -eq $msoFillPicture) `
        'REGRESSION: a texture loaded after ClearTextures still applies'
    Assert (IsBlue (Get-ShapeColour $clearShape 'after_clear')) `
        'and it shows the right image, so the temporary storage came back'

    # Twice, because the first apply is what recreates the directory and the
    # second is what proves it stayed.
    $engine.ClearTextures()
    $again = [long](Field ($engine.LoadTextureBytesAbi($redBytes)) 'texture')
    $engine.ApplyTexture($clearShape, $again)
    Assert (IsRed (Get-ShapeColour $clearShape 'after_clear_twice')) `
        'and the cycle survives being repeated'

    # --- REGRESSION 3 and 22-25: handle states --------------------------------
    $live = [long](Field ($engine.LoadTextureBytesAbi($blueBytes)) 'texture')
    $liveShape = $slide.Shapes.AddShape($msoShapeRectangle, 20, 330, 60, 60)
    $liveShape.Name = 'Live'
    $engine.ApplyTexture($liveShape, $live)
    Assert ($liveShape.Fill.Type -eq $msoFillPicture) 'a live texture handle applies'

    $engine.ReleaseTexture($live)
    $staleMessage = Refused { $engine.ApplyTexture($liveShape, $live) }
    Assert ($staleMessage -ne '') 'a released texture handle is refused'
    Assert ($staleMessage -match 'never recycled|released') `
        'and the refusal says it was released rather than that it never existed'

    $neverMessage = Refused { $engine.ApplyTexture($liveShape, 0x1FFFFFF) }
    Assert ($neverMessage -ne '') 'a handle that was never issued is refused'
    Assert ($neverMessage -match 'never created') 'and is told apart from a released one, which is a different mistake'

    $next = [long](Field ($engine.LoadTextureBytesAbi($redBytes)) 'texture')
    Assert ($next -ne $live -and $next -gt $live) `
        'REGRESSION: the next handle is new, so handles never recycle'

    $releasedImage = [long](Field ($engine.LoadImageAbi($redPath, '')) 'image')
    $engine.ReleaseImageAbi($releasedImage) | Out-Null
    Assert ((Refused { $engine.ImageSizeAbi($releasedImage) }) -ne '') `
        'a released image handle is refused'
    Assert ((Refused { $engine.ReleaseImageAbi($releasedImage) }) -ne '') `
        'and releasing it twice is refused'

    # --- handle-kind safety ---------------------------------------------------
    Assert ((Refused { $engine.ReleaseImageAbi($next) }) -ne '') `
        'a texture handle passed to an image call is refused, not reinterpreted'
    Assert ((Refused { $engine.ApplyTexture($liveShape, $redImage) }) -ne '') `
        'an image handle passed to a texture call is refused'
    Assert ((Refused { $engine.ImageSizeAbi($next) }) -ne '') `
        'and a texture handle is not accepted where an image belongs'

    # --- 27: ClearImages -------------------------------------------------------
    $cleared = $engine.ClearImagesAbi()
    Assert ((Field $cleared 'images') -eq '0') 'ClearImages releases every image'
    Assert ((Refused { $engine.ImageSizeAbi($redImage) }) -ne '') `
        'and an image handle from before the clear is refused afterwards'

    # --- 29: Undo, measured rather than assumed -------------------------------
    <#
    Undo has to be isolated to be measured. Creating a Shape and filling it
    without a boundary lets PowerPoint group both into one entry - the portable
    backend joins whatever entry Office has open - and then Undo removes the
    Shape rather than the fill. A test that reads Fill.Type from the removed
    Shape gets nothing back, which satisfies "not a picture fill" and proves
    nothing at all. So: boundary first, and the Shape is re-acquired by name
    afterwards rather than read through a reference Undo may have invalidated.
    #>
    $undoShape = $slide.Shapes.AddShape($msoShapeRectangle, 100, 330, 60, 60)
    $undoShape.Name = 'UndoTarget'
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()

    $undoTexture = [long](Field ($engine.LoadTextureBytesAbi($blueBytes)) 'texture')
    $engine.ApplyTexture($slide.Shapes.Item('UndoTarget'), $undoTexture)
    Assert ($slide.Shapes.Item('UndoTarget').Fill.Type -eq $msoFillPicture) `
        'the Shape is filled before the Undo'
    $app.StartNewUndoEntry()

    $app.CommandBars.ExecuteMso('Undo')
    $afterUndo = $slide.Shapes | Where-Object { $_.Name -eq 'UndoTarget' } | Select-Object -First 1
    Assert ($null -ne $afterUndo) 'Undo reverted the fill and left the Shape in place'
    if ($afterUndo) {
        Assert ($afterUndo.Fill.Type -ne $msoFillPicture) `
            'and the picture fill is gone, which is what one Undo was meant to revert'
    }
    $app.CommandBars.ExecuteMso('Redo')
    $afterRedo = $slide.Shapes | Where-Object { $_.Name -eq 'UndoTarget' } | Select-Object -First 1
    Assert ($null -ne $afterRedo -and $afterRedo.Fill.Type -eq $msoFillPicture) `
        'and Redo puts the picture back'

    # --- 28: Save, close, reopen ----------------------------------------------
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $null

    $reopened = $app.Presentations.Open($saved, $msoTrue, $msoTrue, $msoTrue)
    try {
        $reopenedSlide = $reopened.Slides.Item(1)
        $reopenedTarget = $reopenedSlide.Shapes.Item('Target')
        Assert ($reopenedTarget.Fill.Type -eq $msoFillPicture) `
            'a reopened Shape still has its picture fill'
        Assert (IsGreen (Get-ShapeColour $reopenedTarget 'reopened_target')) `
            'and the image survived save and reopen'

        Assert (IsBlue (Get-ShapeColour ($reopenedSlide.Shapes.Item('Rewritten')) 'reopened_rewrite')) `
            'the rewritten-path Shape persisted the new image, not the old one'
        Assert (IsGreen (Get-ShapeColour ($reopenedSlide.Shapes.Item('Range2')) 'reopened_range')) `
            'a range-filled Shape persisted its image'
        Assert (($reopenedSlide.Shapes.Item('Quad')).Fill.Type -eq $msoFillPicture) `
            'a quad-warped Shape persisted its fill'
    } finally {
        $reopened.Saved = $true
        $reopened.Close()
    }

    # --- 30: shutdown and re-initialise ---------------------------------------
    # A library that cannot be used after Shutdown would make the add-in
    # single-use, so this asserts the round trip rather than only the call.
    $engine.AbiLifecycle('shutdown') | Out-Null
    $back = $engine.AbiLifecycle('init')
    Assert ((Field $back 'init') -eq '1') 'Shutdown is followed by a successful Init'
    Assert ((Field $back 'textures') -eq '0' -and (Field $back 'images') -eq '0') `
        'and shutdown released everything BlipBridge owned'

    $postShutdown = [long](Field ($engine.LoadTextureBytesAbi($redBytes)) 'texture')
    Assert ($postShutdown -gt $next) `
        'handles still do not recycle across a shutdown and re-init'
} finally {
    if ($presentation) { $presentation.Saved = $true; $presentation.Close() }
    $addin.Connect = $false
    if (-not $KeepArtifacts) { Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue }
}

$results | ForEach-Object { $_ }
if ($failures -gt 0) { throw "$failures check(s) failed" }
"all $($results.Count - 1) checks passed"
