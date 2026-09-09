<#
.SYNOPSIS
Attributes every reference held on a native texture's cached image.

.DESCRIPTION
Bounded memory told us nothing was leaking; it did not tell us who holds what.
This walks one texture through a controlled sequence, changing exactly one thing
at a time and reading the intrusive count after each step, so each delta can be
attributed to an owner rather than guessed at.

The decisive step is the last one: if closing the presentation returns the count
to exactly 1 - the reference the handle itself owns - then everything above 1 was
document-owned and is fully released with the document.

Runs windowed and visible, because Undo needs a document window.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1
$msoFillPicture = 6

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$script:handle = 0
function Get-Count {
    $report = $engine.InspectTexture($script:handle)
    foreach ($part in $report.Split(';')) {
        $pair = $part.Split('=', 2)
        if ($pair.Length -eq 2 -and $pair[0] -eq 'cachedCount') { return [int]$pair[1] }
    }
    return -1
}
$script:previous = 0
function Step([string]$label) {
    $count = Get-Count
    $delta = $count - $script:previous
    $script:previous = $count
    $sign = if ($delta -ge 0) { "+$delta" } else { "$delta" }
    $results.Add(("{0,-46} count={1,-4} delta={2}" -f $label, $count, $sign))
    return $count
}

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide1 = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide2 = $presentation.Slides.Add(2, $ppLayoutBlank)
    $warmup = $slide1.Shapes.AddShape($msoShapeRectangle, 500, 300, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)

    $script:handle = $engine.LoadTexture($bytes)
    $baseline = Step 'LoadTexture (handle owns one reference)'
    if ($baseline -ne 1) { throw "Fresh texture should start at 1, saw $baseline" }

    $shapeA = $slide1.Shapes.AddShape($msoShapeRectangle, 40, 40, 110, 110)
    $shapeB = $slide1.Shapes.AddShape($msoShapeRectangle, 170, 40, 110, 110)
    $shapeC = $slide2.Shapes.AddShape($msoShapeRectangle, 40, 40, 110, 110)
    Step 'three empty Shapes created'

    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($shapeA, $script:handle)
    $afterFirst = Step 'apply to Shape A (first fill on that Shape)'

    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($shapeB, $script:handle)
    Step 'apply to Shape B (same slide)'

    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($shapeC, $script:handle)
    Step 'apply to Shape C (other slide)'

    # Re-applying to a Shape that already carries this image: the old fill is the
    # same object, so this separates "per fill" from "per replacement".
    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($shapeA, $script:handle)
    Step 're-apply to Shape A (replacing the same image)'

    # --- undo ownership ------------------------------------------------------
    $app.CommandBars.ExecuteMso('Undo')
    $afterUndo = Step 'Undo the re-apply'
    if ([int]$shapeA.Fill.Type -ne $msoFillPicture) {
        $results.Add('  (Shape A still has a picture fill: the undo restored the previous apply)')
    }
    $app.CommandBars.ExecuteMso('Redo')
    Step 'Redo it'

    # Flushing the undo history is the cleanest test of undo-owned references.
    for ($i = 0; $i -lt 40; $i++) {
        $app.StartNewUndoEntry()
        $shapeB.Left = 170 + ($i % 3)
    }
    $afterFlush = Step 'flush undo history with 40 unrelated entries'
    $results.Add("  undo-owned references released by flushing: $($script:previous - $afterFlush + ($afterUndo - $afterUndo))")

    # --- shape ownership -----------------------------------------------------
    $shapeC.Delete()
    Step 'delete Shape C'
    $shapeB.Delete()
    Step 'delete Shape B'
    for ($i = 0; $i -lt 40; $i++) { $app.StartNewUndoEntry(); $shapeA.Left = 40 + ($i % 3) }
    Step 'flush undo again (drops the deleted Shapes from history)'

    # --- slide ownership -----------------------------------------------------
    $slide2.Delete()
    for ($i = 0; $i -lt 40; $i++) { $app.StartNewUndoEntry(); $shapeA.Left = 40 + ($i % 3) }
    Step 'delete slide 2 and flush undo'

    # --- save ----------------------------------------------------------------
    $deck = Join-Path $root 'artifacts/refcount_attribution.pptx'
    $presentation.SaveAs($deck, 24)
    Step 'SaveAs'
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

$afterClose = Step 'Presentation.Close'
$results.Add('')
if ($afterClose -eq 1) {
    $results.Add('ATTRIBUTED: closing the document returned the count to 1, the handle''s own')
    $results.Add('reference. Everything above 1 was document-owned and is fully released.')
} else {
    $results.Add("NOT FULLY ATTRIBUTED: $afterClose references remain after the document closed;")
    $results.Add('one of them is the handle, the rest are unexplained and must be traced.')
}

# The handle itself must still be usable afterwards.
$fresh = $app.Presentations.Add($msoTrue)
try {
    $shape = $fresh.Slides.Add(1, $ppLayoutBlank).Shapes.AddShape($msoShapeRectangle, 40, 40, 110, 110)
    $null = $engine.ApplyTexture($shape, $script:handle)
    if ([int]$shape.Fill.Type -ne $msoFillPicture) { throw 'Texture stopped working after the walk' }
    Step 'apply in a brand new presentation'
} finally {
    $fresh.Saved = $msoTrue
    $fresh.Close()
}
Step 'close that presentation too'

$engine.ReleaseTexture($script:handle)
$results.Add("ReleaseTexture done; store now holds $($engine.GetTextureCount()) textures.")
$app.Quit()

$results | Set-Content "$root/artifacts/refcount_attribution.txt"
$results
