<#
.SYNOPSIS
Verifies the multi-Shape apply: what it fills, what it refuses, and what one
Undo does to it.

.DESCRIPTION
One private apply against a ShapeRange fills every member, which is where the
batch speed-up comes from and also where its whole risk lives: the fill reaches
every Shape in the range, so a Connector in the range would reach the private
backend, and a native picture fill on a Connector terminates PowerPoint.

So this asserts, in order:

  filling     every member of a range of supported classes gets the picture
  refusal     one Connector, Line or Table in the range refuses the whole range
  gating      and refuses it *before* the private apply, proved by the entry
              counter rather than by PowerPoint having survived
  undo        that one Undo reverts a whole range fill and one Redo restores it,
              for a uniform range and for a mixed-class range, with a
              single-Shape apply in the same document as the control
  persistence the fills survive save and reopen
  mixed       a range whose members carry different images all end up with the
              one that was applied
  skipping    the per-Shape record is updated for every member, so a later
              ApplyTextureIfChanged on any of them skips
  classes     one range mixing AutoShape, TextBox, WordArt, Freeform and Callout
  groups      a group in the range is filled, and never skipped afterwards
  slides      a ShapeRange cannot span slides, so this path is per slide
  coherence   ApplyPicture, ApplyTexture and ApplyTextureRange all write the
              same record, in every order, so none leaves a stale claim

Everything runs in PowerPoint through the research surface, which drives the
same private path the C ABI does.
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

$results = New-Object System.Collections.Generic.List[string]
$failures = 0
function Assert([bool]$condition, [string]$what) {
    if ($condition) { $script:results.Add("  ok   $what") }
    else { $script:results.Add("  FAIL $what"); $script:failures++ }
}
function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}
function Get-ApplyEntries($engine, $shape) {
    return [int](Get-Field ($engine.ProbeShapePolicy($shape)) 'applyEntries')
}
function Count-Filled($slide, $names) {
    return ($names | ForEach-Object { $slide.Shapes.Item($_).Fill.Type } |
        Where-Object { $_ -eq 6 }).Count
}

$saved = Join-Path ([IO.Path]::GetTempPath()) ("bb_range_" + [Guid]::NewGuid().ToString('N') + '.pptx')


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

    # --- the capability that advertises this API ------------------------------
    # BB_GetCapabilities reports nothing outside PowerPoint by design, so the
    # bits can only be checked from in here. A caller who tests the flag and
    # then finds the call missing has no way to work out which of us is wrong.
    $mask = [uint32](Get-Field ($engine.AbiCapabilities()) 'mask')
    Assert (($mask -band 0x100) -ne 0) `
        ("BB_CAP_RANGE_APPLY is set in host (mask 0x{0:X4})" -f $mask)
    Assert (($mask -band 0x1) -ne 0) 'alongside BB_CAP_NATIVE_BACKEND'

    # --- filling --------------------------------------------------------------
    $names = @()
    for ($i = 0; $i -lt 8; $i++) {
        $shape = $slide.Shapes.AddShape(1, 10 + ($i % 4) * 60, 10 + [math]::Floor($i / 4) * 60, 50, 50)
        $shape.Name = "Member$i"
        $names += $shape.Name
    }
    $range = $slide.Shapes.Range($names)
    Assert ((Count-Filled $slide $names) -eq 0) 'none of the eight Shapes starts with a picture fill'
    $null = $engine.ApplyTextureRange($range, $handleA)
    Assert ((Count-Filled $slide $names) -eq 8) 'one apply against the range fills all eight'

    # --- mixed starting state -------------------------------------------------
    $null = $engine.ApplyTexture($slide.Shapes.Item('Member0'), $handleB)
    $null = $engine.ApplyTexture($slide.Shapes.Item('Member3'), $handleB)
    $null = $engine.ApplyTextureRange($range, $handleA)
    Assert ((Count-Filled $slide $names) -eq 8) 'a range whose members carried different images ends up filled'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member0'), $handleA)) 'skipped') -eq '1') `
        'the range apply recorded what every member now carries, so a later IfChanged skips'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member7'), $handleB)) 'skipped') -eq '0') `
        'and a member asked for a different image is not skipped'

    # --- refusal, before the private apply ------------------------------------
    foreach ($case in @(
            @{ Name = 'Connector'; Make = { $slide.Shapes.AddConnector(1, 10, 300, 200, 350) } },
            @{ Name = 'Line'; Make = { $slide.Shapes.AddLine(10, 380, 200, 420) } },
            @{ Name = 'Table'; Make = { $slide.Shapes.AddTable(2, 2, 250, 250, 200, 80) } },
            @{ Name = 'Chart'; Make = { $slide.Shapes.AddChart2(-1, 51, 250, 150, 200, 120) } })) {
        $extra = & $case.Make
        $extra.Name = "Bad$($case.Name)"
        $mixed = $slide.Shapes.Range($names + $extra.Name)

        $before = Get-ApplyEntries $engine $slide.Shapes.Item('Member0')
        $refused = $false
        $message = ''
        try { $null = $engine.ApplyTextureRange($mixed, $handleB) }
        catch { $refused = $true; $message = $_.Exception.Message }
        Assert $refused "a range containing a $($case.Name) is refused"
        Assert ($message -match 'Range member') 'and the message names which member'
        Assert ((Get-ApplyEntries $engine $slide.Shapes.Item('Member0')) -eq $before) `
            'and nothing entered the private apply'
        $extra.Delete()
    }

    # The good members must be untouched by a refusal, not half-filled.
    Assert ((Count-Filled $slide $names) -eq 8) 'the refusals left the other members as they were'

    # --- undo and redo --------------------------------------------------------
    $undoNames = @()
    for ($i = 0; $i -lt 4; $i++) {
        $shape = $slide.Shapes.AddShape(1, 300 + $i * 60, 300, 50, 50)
        $shape.Name = "Undo$i"
        $undoNames += $shape.Name
    }
    $undoRange = $slide.Shapes.Range($undoNames)
    Assert ((Count-Filled $slide $undoNames) -eq 0) 'four fresh Shapes, none filled'
    $null = $engine.ApplyTextureRange($undoRange, $handleA)
    Assert ((Count-Filled $slide $undoNames) -eq 4) 'the range apply fills all four'

    <#
    One range apply is one undo entry, the same as Office's own range fill.

    Counted by putting a marker Shape on the undo stack first and undoing until
    the marker goes, which measures the entries an operation added rather than
    guessing from what Undo appears to do:

        operation                                  4 members  8  16
        nothing                                            0  0   0
        N x BB_ApplyTexture                                4  8  16
        BB_ApplyTextureRange                               1  1   1
        Office's own ShapeRange.Fill.UserPicture           1  1   1

    So the range path matches the public Office operation exactly, and the
    per-Shape path is the one that leaves an entry per Shape - which it always
    has. The check below covers a uniform range and the mixed-class range, so
    the "one entry" claim is not resting on a single Shape type.
    #>
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()
    $undoCount = 0
    for ($u = 1; $u -le 12; $u++) {
        $app.CommandBars.ExecuteMso('Undo')
        if ((Count-Filled $slide $undoNames) -eq 0) { $undoCount = $u; break }
    }
    Assert ($undoCount -eq 1) `
        "one Undo reverts the whole range fill (took $undoCount)"
    $app.CommandBars.ExecuteMso('Redo')
    Assert ((Count-Filled $slide $undoNames) -eq 4) 'and one Redo restores all four'

    # The control, in the same document and the same undo stack: a single-Shape
    # native apply is undone by one Undo, which is what our per-Shape path has
    # always done.
    $control = $slide.Shapes.AddShape(1, 300, 380, 50, 50)
    $control.Name = 'UndoControl'
    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($control, $handleA)
    Assert ($control.Fill.Type -eq 6) 'a single-Shape apply fills its Shape'
    $app.CommandBars.ExecuteMso('Undo')
    Assert ($control.Fill.Type -ne 6) 'and one Undo reverts it, so the undo stack is working'
    $control.Delete()

    # --- mixed classes --------------------------------------------------------
    # Every class with a validated native path, in one range. WordArt and the
    # Freeform are the ones worth having here: WordArt reports msoAutoShape, the
    # same type a Connector reports, and a Freeform reports a ShapeRange as its
    # parent - both have caught something before.
    $mixedNames = @()
    $auto = $slide.Shapes.AddShape(1, 20, 440, 50, 50); $auto.Name = 'MixAuto'
    $mixedNames += $auto.Name
    $box = $slide.Shapes.AddTextbox(1, 90, 440, 60, 40); $box.Name = 'MixBox'
    $mixedNames += $box.Name
    $art = $slide.Shapes.AddTextEffect(0, 'Bb', 'Arial', 20, $msoFalse, $msoFalse, 170, 440)
    $art.Name = 'MixArt'; $mixedNames += $art.Name
    $builder = $slide.Shapes.BuildFreeform(1, 260, 440)
    $null = $builder.AddNodes(0, 0, 320, 440)
    $null = $builder.AddNodes(0, 0, 320, 490)
    $null = $builder.AddNodes(0, 0, 260, 440)
    $curve = $builder.ConvertToShape(); $curve.Name = 'MixFree'
    $mixedNames += $curve.Name
    $callout = $slide.Shapes.AddShape(106, 340, 440, 50, 50); $callout.Name = 'MixCallout'
    $mixedNames += $callout.Name

    $null = $engine.ApplyTextureRange($slide.Shapes.Range($mixedNames), $handleA)
    Assert ((Count-Filled $slide $mixedNames) -eq $mixedNames.Count) `
        "a range mixing AutoShape, TextBox, WordArt, Freeform and Callout fills all $($mixedNames.Count)"

    # Alternating images through the range, which is what a caller animating a
    # whole slide would do.
    $null = $engine.ApplyTextureRange($slide.Shapes.Range($mixedNames), $handleB)
    Assert ((Count-Filled $slide $mixedNames) -eq $mixedNames.Count) 'and again with a second image'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('MixArt'), $handleB)) 'skipped') -eq '1') `
        'the second image is what the record names'

    # One entry for a range of five different classes too, so the claim is not
    # resting on a range of identical AutoShapes. Counted with a marker Shape
    # rather than by watching the fills: these are filled already, so Fill.Type
    # cannot tell one image from another.
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()
    $mixedMarker = $slide.Shapes.AddShape(1, 480, 200, 40, 40)
    $mixedMarker.Name = 'MixMarker'
    $app.StartNewUndoEntry()
    $null = $engine.ApplyTextureRange($slide.Shapes.Range($mixedNames), $handleA)
    $mixedUndo = -1
    for ($u = 1; $u -le 12; $u++) {
        $app.CommandBars.ExecuteMso('Undo')
        $markerGone = $true
        try { $null = $slide.Shapes.Item('MixMarker'); $markerGone = $false } catch { }
        if ($markerGone) { $mixedUndo = $u - 1; break }
    }
    Assert ($mixedUndo -eq 1) `
        "a range of five different classes is also one undo entry (counted $mixedUndo)"
    for ($u = 1; $u -le 2; $u++) { $app.CommandBars.ExecuteMso('Redo') }
    Assert ((Count-Filled $slide $mixedNames) -eq $mixedNames.Count) 'and Redo restores them'

    # Undo and Redo change fills without telling BlipBridge, exactly like any
    # other change made behind its back, so the record can be stale afterwards.
    # InvalidateShape is the documented remedy and is what this asserts.
    $null = $engine.InvalidateShape($slide.Shapes.Item('MixArt'))
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('MixArt'), $handleA)) 'skipped') -eq '0') `
        'InvalidateShape clears a record left stale by Undo'

    # --- a group in the range -------------------------------------------------
    # msoGroup has a validated native path, so a group is filled rather than
    # refused - and filling one changes what its children render, which is why
    # neither a group nor anything inside one is ever keyed for skipping.
    $g1 = $slide.Shapes.AddShape(1, 420, 440, 40, 40); $g1.Name = 'G1'
    $g2 = $slide.Shapes.AddShape(1, 470, 440, 40, 40); $g2.Name = 'G2'
    $group = $slide.Shapes.Range(@('G1', 'G2')).Group()
    $group.Name = 'MixGroup'
    $withGroup = $engine.ApplyTextureRange($slide.Shapes.Range(@('MixAuto', 'MixGroup')), $handleA)
    Assert ($null -ne $withGroup) 'a range containing a group is accepted, not refused'
    Assert ($slide.Shapes.Item('MixGroup').Fill.Type -eq 6) 'and the group gets the picture fill'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('MixGroup'), $handleA)) 'skipped') -eq '0') `
        'a group is never skipped, because its fill and its children are entangled'
    $slide.Shapes.Item('MixGroup').Delete()

    # --- a range cannot span slides -------------------------------------------
    # Shapes.Range is a member of one Slide's Shapes collection, so there is no
    # cross-slide ShapeRange to hand this path. A caller filling several slides
    # makes one call per slide, and the per-call fixed cost is paid per slide.
    $second = $presentation.Slides.Add($presentation.Slides.Count + 1, $ppLayoutBlank)
    $far = $second.Shapes.AddShape(1, 20, 20, 50, 50); $far.Name = 'FarShape'
    $crossFailed = $false
    try { $null = $slide.Shapes.Range(@('MixAuto', 'FarShape')) } catch { $crossFailed = $true }
    Assert $crossFailed 'a ShapeRange cannot be built across two slides, so this path is per slide'
    $second.Delete()

    # --- a Shape is not a ShapeRange ------------------------------------------
    # Passing one must say so rather than being quietly reinterpreted as a range
    # of one: the two APIs mean different things and a caller who mixed them up
    # needs to be told which one they wanted.
    $lone = $slide.Shapes.Item('Member0')
    $refusedShape = $false
    $shapeMessage = ''
    try { $null = $engine.ApplyTextureRange($lone, $handleA) }
    catch { $refusedShape = $true; $shapeMessage = $_.Exception.Message }
    Assert $refusedShape 'a Shape passed where a ShapeRange belongs is refused'
    Assert ($shapeMessage -match 'ShapeRange') 'and the message says it wanted a ShapeRange'
    Assert ($shapeMessage -match 'BB_ApplyTexture') 'and names the call that does take a Shape'

    # --- cache coherence ------------------------------------------------------
    # Every API that writes a fill must update the same record, or one of them
    # will skip on the strength of an image another one replaced. Each sequence
    # below ends by asking for image A through the path that skips; the fill has
    # to end up as A, and the middle call must not have left a claim that lets
    # the last one do nothing.
    $coherent = $slide.Shapes.AddShape(1, 420, 200, 60, 60)
    $coherent.Name = 'Coherent'
    $coherentRange = $slide.Shapes.Range(@('Coherent'))

    <#
    Each image has to be asked about through the API that owns it. A handle from
    LoadTexture and the image ApplyPicture decodes from the same file are two
    different images with two different internal ids - same picture, same bytes,
    separate resources - so asking ApplyTextureIfChanged whether a Shape carries
    "the file" would be asking the wrong question and would fail for a reason
    that is not a bug.

    So the closing ApplyPicture A is checked through the picture cache's own skip
    counter: repeat it, and it must skip, which it can only do if the record
    correctly says the Shape carries that image.
    #>
    function Get-Skipped($engine) {
        return [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    }

    foreach ($sequence in @(
            @{ Name = 'ApplyPicture A -> Range B -> ApplyPicture A';
               Steps = { $null = $engine.ApplyPicture($coherent, $textureA)
                         $null = $engine.ApplyTextureRange($coherentRange, $handleB)
                         $null = $engine.ApplyPicture($coherent, $textureA) } },
            @{ Name = 'ApplyTexture A -> Range B -> ApplyPicture A';
               Steps = { $null = $engine.ApplyTexture($coherent, $handleA)
                         $null = $engine.ApplyTextureRange($coherentRange, $handleB)
                         $null = $engine.ApplyPicture($coherent, $textureA) } },
            @{ Name = 'Range A -> ApplyTexture B -> ApplyPicture A';
               Steps = { $null = $engine.ApplyTextureRange($coherentRange, $handleA)
                         $null = $engine.ApplyTexture($coherent, $handleB)
                         $null = $engine.ApplyPicture($coherent, $textureA) } })) {
        & $sequence.Steps
        Assert ($coherent.Fill.Type -eq 6) "$($sequence.Name): ends with a picture fill"

        # The record says A: repeating A skips.
        $before = Get-Skipped $engine
        $null = $engine.ApplyPicture($coherent, $textureA)
        Assert ((Get-Skipped $engine) -eq ($before + 1)) `
            "$($sequence.Name): the record names A, so repeating A skips"

        # And it does not say B: the middle call's image left no stale claim.
        Assert ((Get-Field ($engine.ApplyTextureIfChanged($coherent, $handleB)) 'skipped') -eq '0') `
            "$($sequence.Name): and B is not skipped on a stale claim"
    }

    # And the range's own record, over more than one member: a range apply must
    # leave every member claiming what the range put on it.
    $null = $engine.ApplyTextureRange($slide.Shapes.Range($names), $handleA)
    $null = $engine.ApplyPicture($slide.Shapes.Item('Member2'), $textureB)
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member2'), $handleA)) 'skipped') -eq '0') `
        'ApplyPicture on one member clears that member from the range apply record'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member3'), $handleA)) 'skipped') -eq '1') `
        'and leaves the other members alone'
    $coherent.Delete()

    # --- persistence ----------------------------------------------------------
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopenedSlide = $presentation.Slides.Item(1)
    $survived = ($names + $undoNames | ForEach-Object { $reopenedSlide.Shapes.Item($_).Fill.Type } |
        Where-Object { $_ -eq 6 }).Count
    Assert ($survived -eq 12) "every range-applied fill survives save and reopen ($survived of 12)"
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
