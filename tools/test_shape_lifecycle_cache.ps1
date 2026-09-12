<#
.SYNOPSIS
Shape lifecycle against the per-Shape "already applied" cache.

.DESCRIPTION
The skip cache remembers which texture each Shape last received, keyed by
presentation identity plus SlideID plus Shape.Id. Every operation that creates,
moves, regroups or resurrects a Shape is a chance for that key to mean something
it did not mean before - and the failure mode is silent: a Shape shows the wrong
image because an apply was skipped that should not have been.

So each operation below is checked for the one thing that actually matters:

    was an apply skipped when real work was needed?

The skip counter makes that observable. A correct skip increments it; an
incorrect one would too, which is why every case asserts the *expected* delta
rather than merely that the call returned.

Covered: Duplicate, Copy/Paste, Group, Ungroup, move to another slide, delete,
undo delete, redo delete, and presentation reopen.
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
$ppLayoutBlank = 12
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


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
function Skipped { [int](Get-Field ($script:engine.PictureCacheStats()) 'skipped') }

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $second = $presentation.Slides.Add(2, $ppLayoutBlank)
    $null = $engine.ClearPictureCache()

    # --- Duplicate ------------------------------------------------------------
    $original = $slide.Shapes.AddShape(1, 20, 20, 100, 100)
    $null = $engine.ApplyPicture($original, $textureA)
    $copy = $original.Duplicate()
    $copy.Left = 140
    # The duplicate inherits the picture fill, but it is a different Shape with a
    # different Id, so nothing should be remembered about it yet.
    $before = Skipped
    $null = $engine.ApplyPicture($copy, $textureA)
    Assert ((Skipped) -eq $before) 'a duplicated Shape is not skipped on its first apply'
    # And the second time it should be.
    $before = Skipped
    $null = $engine.ApplyPicture($copy, $textureA)
    Assert ((Skipped) -eq ($before + 1)) 'the duplicate is skipped on its second apply'
    # The original must be unaffected by anything the duplicate did.
    $before = Skipped
    $null = $engine.ApplyPicture($original, $textureA)
    Assert ((Skipped) -eq ($before + 1)) 'the original keeps its own cache entry'
    # A different image on the duplicate must still do real work.
    $before = Skipped
    $null = $engine.ApplyPicture($copy, $textureB)
    Assert ((Skipped) -eq $before) 'a different image on the duplicate is not skipped'

    # --- Copy / Paste ---------------------------------------------------------
    $original.Copy()
    $pasted = $slide.Shapes.Paste()
    $pasted.Left = 260
    $pasted.Top = 20
    $before = Skipped
    $null = $engine.ApplyPicture($pasted, $textureA)
    Assert ((Skipped) -eq $before) 'a pasted Shape is not skipped on its first apply'
    Assert ($pasted.Fill.Type -eq 6) 'and it ends up with the picture fill'

    # --- Move to another slide ------------------------------------------------
    $traveller = $slide.Shapes.AddShape(1, 20, 160, 80, 80)
    $null = $engine.ApplyPicture($traveller, $textureA)
    $null = $engine.ApplyPicture($traveller, $textureA)   # now cached
    $traveller.Cut()
    $moved = $second.Shapes.Paste()
    # Different slide means a different SlideID, so the old key cannot match.
    $before = Skipped
    $null = $engine.ApplyPicture($moved, $textureA)
    Assert ((Skipped) -eq $before) 'a Shape moved to another slide is not skipped'

    # --- Group ----------------------------------------------------------------
    $a = $slide.Shapes.AddShape(1, 400, 20, 60, 60)
    $b = $slide.Shapes.AddShape(1, 400, 100, 60, 60)
    $null = $engine.ApplyPicture($a, $textureA)
    $null = $engine.ApplyPicture($a, $textureA)          # cached while top-level
    $group = $slide.Shapes.Range(@($a.Name, $b.Name)).Group()
    $child = $group.GroupItems.Item(1)
    <#
     Neither a group nor anything inside one is cached. Filling a group changes
     what its children render - measured in probe_group_fill_propagation.ps1,
     33,515 bytes against 35,725 - so a child's remembered texture goes stale the
     moment its group is filled, and the reverse holds too. A skip there would
     leave the wrong picture on screen silently, so both sides always do real
     work.
    #>
    $before = Skipped
    $null = $engine.ApplyPicture($child, $textureA)
    $null = $engine.ApplyPicture($child, $textureA)
    Assert ((Skipped) -eq $before) 'a group child is never skipped'
    Assert ($child.Fill.Type -eq 6) 'and it is filled correctly every time'

    $before = Skipped
    $null = $engine.ApplyPicture($group, $textureA)
    $null = $engine.ApplyPicture($group, $textureA)
    Assert ((Skipped) -eq $before) 'a group is never skipped either'
    Assert ($group.Fill.Type -eq 6) 'and the group is filled'

    # --- Ungroup --------------------------------------------------------------
    # Once released, a Shape is top-level again and becomes cacheable.
    $released = $group.Ungroup()
    $first = $released.Item(1)
    $before = Skipped
    $null = $engine.ApplyPicture($first, $textureB)
    Assert ((Skipped) -eq $before) 'an ungrouped child does real work for a new image'
    Assert ($first.Fill.Type -eq 6) 'and is filled'
    $before = Skipped
    $null = $engine.ApplyPicture($first, $textureB)
    Assert ((Skipped) -eq ($before + 1)) 'and is cacheable again once out of the group'

    # --- Delete, undo, redo ---------------------------------------------------
    $doomed = $slide.Shapes.AddShape(1, 520, 20, 70, 70)
    $null = $engine.ApplyPicture($doomed, $textureA)
    $null = $engine.ApplyPicture($doomed, $textureA)      # cached
    $doomedId = $doomed.Id
    <#
     Close the entry holding the Shape and its fill, so the delete below is an
     entry of its own and the three Undo/Redo steps that follow act on exactly
     the delete. Without this the grouping is PowerPoint's to choose - it may
     put the creation, the fill and the delete in one entry - and an Undo would
     unwind past the slide, leaving the rest of this suite operating on a dead
     object for a reason that has nothing to do with the cache it is testing.
    #>
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()
    $doomed.Delete()
    $app.StartNewUndoEntry()

    <#
     Undo comes first, with nothing in between. An apply performed after the
     delete would be the top of the undo stack, and Undo would reverse *that*
     instead - which is what made this leg silently skip itself the first time.
    #>
    $app.CommandBars.ExecuteMso('Undo')
    $restored = $null
    foreach ($candidate in $slide.Shapes) {
        if ($candidate.Id -eq $doomedId) { $restored = $candidate }
    }
    Assert ($null -ne $restored) 'undo brings the deleted Shape back'

    <#
     Redo before any new apply. Performing one would clear the redo stack, which
     is ordinary PowerPoint behaviour and would make ExecuteMso('Redo') fail for
     a reason that has nothing to do with this library.
    #>
    $app.CommandBars.ExecuteMso('Redo')
    $goneAgain = $true
    foreach ($candidate in $slide.Shapes) {
        if ($candidate.Id -eq $doomedId) { $goneAgain = $false }
    }
    Assert $goneAgain 'redo deletes it again'
    $app.CommandBars.ExecuteMso('Undo')
    $restored = $null
    foreach ($candidate in $slide.Shapes) {
        if ($candidate.Id -eq $doomedId) { $restored = $candidate }
    }
    Assert ($null -ne $restored) 'and a second undo restores it once more'

    if ($restored) {
        # It returns with its fill intact and its old Id, so its cache entry still
        # matches. That is correct - nothing about it changed - and what must not
        # happen is a wrong image or a failure.
        $null = $engine.ApplyPicture($restored, $textureB)
        Assert ($restored.Fill.Type -eq 6) 'an undeleted Shape still applies correctly'
        $before = Skipped
        $null = $engine.ApplyPicture($restored, $textureB)
        Assert ((Skipped) -eq ($before + 1)) 'and caches normally afterwards'
    }

    # Nothing may reach through a deleted Shape: no receiver is retained, and the
    # cache holds identity values rather than pointers, so there is nothing to
    # dangle. A fresh Shape on the same slide must behave as new.
    $replacement = $slide.Shapes.AddShape(1, 520, 120, 70, 70)
    $before = Skipped
    $null = $engine.ApplyPicture($replacement, $textureA)
    Assert ((Skipped) -eq $before) 'a Shape created after a delete is not skipped'
    Assert ($replacement.Id -ne $doomedId) 'and PowerPoint did not reuse the deleted Id'

    # --- Presentation reopen --------------------------------------------------
    $saved = Join-Path $env:TEMP ('bblife_' + [guid]::NewGuid().ToString('N') + '.pptx')
    $presentation.SaveAs($saved)
    $anchor = $app.Presentations.Add($msoTrue)
    $presentation.Close()
    $presentation = $null
    $reopened = $app.Presentations.Open($saved, $false, $false, $msoTrue)
    <#
     A reopened document can hand out the same SlideID and Shape.Id as the one
     just closed, so old entries could match new Shapes. That is the documented
     reason ClearPictureCache exists, and this asserts the remedy works rather
     than pretending the hazard is absent.
    #>
    $null = $engine.ClearPictureCache()
    $reopenedShape = $reopened.Slides.Item(1).Shapes.Item(1)
    $before = Skipped
    $null = $engine.ApplyPicture($reopenedShape, $textureA)
    Assert ((Skipped) -eq $before) 'after ClearPictureCache a reopened Shape is not skipped'
    Assert ($reopenedShape.Fill.Type -eq 6) 'and it is filled'
    $reopened.Saved = $msoTrue
    $reopened.Close()
    $presentation = $anchor

    # --- nothing was retained --------------------------------------------------
    $stats = $engine.PictureCacheStats()
    $results.Add("  note final cache: $stats")
    $probe = $presentation.Slides.Add(1, $ppLayoutBlank).Shapes.AddShape(1, 10, 10, 40, 40)
    $probe.Fill.UserPicture($textureA)
    Assert ($probe.Fill.Type -eq 6) 'ordinary Fill.UserPicture still works at the end'
    Remove-Item $saved -ErrorAction SilentlyContinue
} finally {
    if ($presentation) { try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { } }
}

$results | Set-Content "$root/artifacts/shape_lifecycle_cache.txt"
$results
if ($failures -gt 0) { throw "$failures lifecycle cache assertions failed" }
'Shape lifecycle cache: all assertions passed.'
