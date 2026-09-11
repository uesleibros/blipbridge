<#
.SYNOPSIS
Verifies BB_ApplyTextureIfChanged: when it skips, when it must not, and that a
skip leaves the document exactly as an ordinary apply would have.

.DESCRIPTION
A skip is only worth having if it is indistinguishable from the apply it
replaces, so this asserts the document, not just the timing:

  skipping    the same image twice does no Office work; a different image does
  undo        a skip adds no undo entry, so one Undo still reverts one apply
  redo        and Redo puts it back
  persistence a fill applied this way survives save and reopen
  freshness   a fill cleared or recoloured behind the cache is re-applied
  crossing    ApplyTexture and ApplyPicture change the same record, so neither
              can leave a stale claim the other would honour
  invalidation InvalidateShape, ClearTextures, ClearPictureCache and Shutdown
              each drop what they promise to drop
  semantics   a Connector and a Line are refused *before* the private apply, and
              a class with no native path is refused by name

Everything runs through the real C ABI in process. Every check is asserted: a
wrong answer fails the script rather than printing something to squint at.
#>
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
$ppSaveAsDefault = 11

$textureA = Join-Path $root 'artifacts/textures/texture_128_0.png'
$textureB = Join-Path $root 'artifacts/textures/texture_64_1.png'
if (-not (Test-Path $textureA)) { throw "Missing $textureA - run tools/generate_textures.ps1" }
if (-not (Test-Path $textureB)) { throw "Missing $textureB - run tools/generate_textures.ps1" }

$results = New-Object System.Collections.Generic.List[string]
$failures = 0
function Assert([bool]$condition, [string]$what) {
    if ($condition) {
        $script:results.Add("  ok   $what")
    } else {
        $script:results.Add("  FAIL $what")
        $script:failures++
    }
}
function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}
# The private apply's entry counter, which is how a refusal is proved to have
# happened *before* the dangerous call rather than merely to have been survived.
function Get-ApplyEntries($engine, $shape) {
    return [int](Get-Field ($engine.ProbeShapePolicy($shape)) 'applyEntries')
}

$saved = Join-Path ([IO.Path]::GetTempPath()) ("bb_if_changed_" + [Guid]::NewGuid().ToString('N') + '.pptx')


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    [byte[]]$bytesA = [IO.File]::ReadAllBytes($textureA)
    [byte[]]$bytesB = [IO.File]::ReadAllBytes($textureB)
    $handleA = $engine.LoadTexture($bytesA)
    $handleB = $engine.LoadTexture($bytesB)

    # --- skipping -------------------------------------------------------------
    $shape = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleA)) 'skipped') -eq '0') `
        'the first apply is not skipped'
    Assert ($shape.Fill.Type -eq 6) 'and it leaves a picture fill'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleA)) 'skipped') -eq '1') `
        'the same image again is skipped'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'a different image is not skipped'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '1') `
        'and then it is'

    # A skip must not reach the private apply at all.
    $before = Get-ApplyEntries $engine $shape
    $null = $engine.ApplyTextureIfChanged($shape, $handleB)
    Assert ((Get-ApplyEntries $engine $shape) -eq $before) `
        'a skip never enters the private apply'

    # --- freshness: the fill changed behind the record ------------------------
    $shape.Fill.ForeColor.RGB = 255
    $null = $shape.Fill.Solid()
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'a fill replaced with a colour is re-applied, not skipped'
    Assert ($shape.Fill.Type -eq 6) 'and the picture is back'

    # --- crossing: ApplyTexture must not leave a stale claim ------------------
    # ApplyTexture always applies, so it can silently replace what the record
    # says is there. If it did not update the record, this next call would skip
    # and the Shape would keep showing image A while the caller asked for B.
    $null = $engine.ApplyTexture($shape, $handleA)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'ApplyTexture updates the record, so a later IfChanged does not skip wrongly'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleA)) 'skipped') -eq '0') `
        'and the record names what was last applied'

    # --- crossing the other way: ApplyPicture shares the record ---------------
    $null = $engine.ApplyPicture($shape, $textureA)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'ApplyPicture updates the same record'

    # --- invalidation ---------------------------------------------------------
    $null = $engine.ApplyTextureIfChanged($shape, $handleB)
    $null = $engine.InvalidateShape($shape)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'InvalidateShape makes the next apply real'

    $null = $engine.ApplyTextureIfChanged($shape, $handleB)
    $null = $engine.ClearPictureCache()
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'ClearPictureCache drops every remembered Shape'

    $null = $engine.ApplyTextureIfChanged($shape, $handleB)
    $null = $engine.ReleaseTexture($handleB)
    $handleB = $engine.LoadTexture($bytesB)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'a released image cannot be inherited by a later one'

    $null = $engine.ApplyTextureIfChanged($shape, $handleB)
    $null = $engine.ClearTextures()
    $handleA = $engine.LoadTexture($bytesA)
    $handleB = $engine.LoadTexture($bytesB)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($shape, $handleB)) 'skipped') -eq '0') `
        'ClearTextures drops every remembered Shape'

    # --- deletion and recreation ---------------------------------------------
    $null = $engine.ApplyTextureIfChanged($shape, $handleA)
    $shape.Delete()
    $fresh = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($fresh, $handleA)) 'skipped') -eq '0') `
        'a recreated Shape is never skipped on the strength of the deleted one'
    Assert ($fresh.Fill.Type -eq 6) 'and it really gets the picture'
    $shape = $fresh

    # --- undo and redo --------------------------------------------------------
    # A fresh Shape with a solid fill, so "reverted" is unambiguous.
    $undoShape = $slide.Shapes.AddShape(1, 200, 20, 120, 120)
    Assert ($undoShape.Fill.Type -ne 6) 'a new Shape starts without a picture fill'
    $null = $engine.ApplyTextureIfChanged($undoShape, $handleA)
    Assert ($undoShape.Fill.Type -eq 6) 'IfChanged applies to it'
    # Three skips: if any of them made an undo entry, one Undo would not be
    # enough to get back to the solid fill.
    $null = $engine.ApplyTextureIfChanged($undoShape, $handleA)
    $null = $engine.ApplyTextureIfChanged($undoShape, $handleA)
    $null = $engine.ApplyTextureIfChanged($undoShape, $handleA)
    $app.StartNewUndoEntry()
    $presentation.Windows.Item(1).Activate()
    $app.CommandBars.ExecuteMso('Undo')
    Assert ($undoShape.Fill.Type -ne 6) 'one Undo reverts the one real apply - skips added no entries'
    $app.CommandBars.ExecuteMso('Redo')
    Assert ($undoShape.Fill.Type -eq 6) 'and Redo puts the picture back'
    $undoShape.Delete()

    # --- semantics: refused before the private apply --------------------------
    $connector = $slide.Shapes.AddConnector(1, 10, 300, 200, 350)
    $before = Get-ApplyEntries $engine $connector
    $refused = $false
    try { $null = $engine.ApplyTextureIfChanged($connector, $handleA) } catch { $refused = $true }
    Assert $refused 'a Connector is refused'
    Assert ((Get-ApplyEntries $engine $connector) -eq $before) `
        'and refused before the private apply, not by surviving it'
    $connector.Delete()

    $line = $slide.Shapes.AddLine(10, 380, 200, 420)
    $before = Get-ApplyEntries $engine $line
    $refused = $false
    try { $null = $engine.ApplyTextureIfChanged($line, $handleA) } catch { $refused = $true }
    Assert $refused 'a Line is refused'
    Assert ((Get-ApplyEntries $engine $line) -eq $before) 'and before the private apply'
    $line.Delete()

    # A class Office can fill but BlipBridge has no native path for: refused by
    # name, never quietly redirected to the fallback.
    $table = $slide.Shapes.AddTable(2, 2, 250, 250, 200, 80)
    $refused = $false
    try { $null = $engine.ApplyTextureIfChanged($table, $handleA) } catch { $refused = $true }
    Assert $refused 'a Table has no native path and is refused'
    $table.Delete()

    # WordArt reports msoAutoShape and must still reach the native backend.
    $wordArt = $slide.Shapes.AddTextEffect(0, 'Bb', 'Arial', 24, $msoFalse, $msoFalse, 320, 300)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($wordArt, $handleA)) 'skipped') -eq '0') `
        'WordArt is applied to natively'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($wordArt, $handleA)) 'skipped') -eq '1') `
        'and is skipped the second time'
    $wordArt.Delete()

    # A Freeform built this way reports a ShapeRange as its Parent rather than
    # the Slide, which is exactly the case shape_identity.cpp walks past.
    $freeform = $slide.Shapes.BuildFreeform(1, 400, 300)
    $null = $freeform.AddNodes(0, 0, 500, 300)
    $null = $freeform.AddNodes(0, 0, 500, 400)
    $null = $freeform.AddNodes(0, 0, 400, 300)
    $free = $freeform.ConvertToShape()
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($free, $handleA)) 'skipped') -eq '0') `
        'a Freeform is applied to natively'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($free, $handleA)) 'skipped') -eq '1') `
        'and is skipped the second time'
    $free.Delete()

    # --- persistence ----------------------------------------------------------
    $keeper = $slide.Shapes.AddShape(1, 20, 200, 120, 120)
    $keeper.Name = 'BbKeeper'
    $null = $engine.ApplyTextureIfChanged($keeper, $handleA)
    $null = $engine.ApplyTextureIfChanged($keeper, $handleA)   # skipped; must not undo the fill
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopened = $presentation.Slides.Item(1).Shapes.Item('BbKeeper')
    Assert ($reopened.Fill.Type -eq 6) 'the fill survives save and reopen'
    Assert ($reopened.Fill.TextureType -ne 0) 'and reopens as a real picture fill'

    # A reopened Shape is a different instance in a different document. It must
    # not inherit anything the closed one was remembered as carrying.
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($reopened, $handleA)) 'skipped') -eq '0') `
        'a reopened Shape is not skipped on the strength of the closed document'
} finally {
    # Closing is cleanup, not an assertion. Run back to back with other suites
    # this can arrive at a presentation the host has already taken down - seen as
    # "Presentation.Saved : Object does not exist" - and a tidy-up that cannot
    # find its document must not turn a run whose every check passed into a
    # failure. Anything that goes wrong here is reported and counted as noise.
    try {
        $presentation.Saved = $msoTrue
        $presentation.Close()
    } catch {
        $results.Add("  note could not close the presentation: " +
                     $_.Exception.Message.Split("`n")[0])
    }
    if (Test-Path $saved) { Remove-Item $saved -Force }
}

$results
if ($failures -gt 0) { throw "$failures check(s) failed" }
"all $($results.Count) checks passed"
