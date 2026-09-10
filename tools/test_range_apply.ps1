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
  undo        that the range apply makes NO undo entry, with a single-Shape
              apply in the same document as the control that proves the undo
              stack was working - this is the finding that decides the path
              cannot be what ApplyTexture does
  persistence the fills survive save and reopen
  mixed       a range whose members carry different images all end up with the
              one that was applied
  skipping    the per-Shape record is updated for every member, so a later
              ApplyTextureIfChanged on any of them skips

Everything runs in PowerPoint through the research surface, which drives the
same private path the C ABI does.
#>
$ErrorActionPreference = 'Stop'
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

$app = New-Object -ComObject PowerPoint.Application
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

    # --- filling --------------------------------------------------------------
    $names = @()
    for ($i = 0; $i -lt 8; $i++) {
        $shape = $slide.Shapes.AddShape(1, 10 + ($i % 4) * 60, 10 + [math]::Floor($i / 4) * 60, 50, 50)
        $shape.Name = "Member$i"
        $names += $shape.Name
    }
    $range = $slide.Shapes.Range($names)
    Assert ((Count-Filled $slide $names) -eq 0) 'none of the eight Shapes starts with a picture fill'
    $null = $engine.ApplyTextureToRange($range, $handleA, 1)
    Assert ((Count-Filled $slide $names) -eq 8) 'one apply against the range fills all eight'

    # --- mixed starting state -------------------------------------------------
    $null = $engine.ApplyTexture($slide.Shapes.Item('Member0'), $handleB)
    $null = $engine.ApplyTexture($slide.Shapes.Item('Member3'), $handleB)
    $null = $engine.ApplyTextureToRange($range, $handleA, 1)
    Assert ((Count-Filled $slide $names) -eq 8) 'a range whose members carried different images ends up filled'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member0'), $handleA)) 'skipped') -eq '1') `
        'the range apply recorded what every member now carries, so a later IfChanged skips'
    Assert ((Get-Field ($engine.ApplyTextureIfChanged($slide.Shapes.Item('Member7'), $handleB)) 'skipped') -eq '0') `
        'and a member asked for a different image is not skipped'

    # --- refusal, before the private apply ------------------------------------
    foreach ($case in @(
            @{ Name = 'Connector'; Make = { $slide.Shapes.AddConnector(1, 10, 300, 200, 350) } },
            @{ Name = 'Line'; Make = { $slide.Shapes.AddLine(10, 380, 200, 420) } },
            @{ Name = 'Table'; Make = { $slide.Shapes.AddTable(2, 2, 250, 250, 200, 80) } })) {
        $extra = & $case.Make
        $extra.Name = "Bad$($case.Name)"
        $mixed = $slide.Shapes.Range($names + $extra.Name)

        $before = Get-ApplyEntries $engine $slide.Shapes.Item('Member0')
        $refused = $false
        $message = ''
        try { $null = $engine.ApplyTextureToRange($mixed, $handleB, 1) }
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
    $null = $engine.ApplyTextureToRange($undoRange, $handleA, 1)
    Assert ((Count-Filled $slide $undoNames) -eq 4) 'the range apply fills all four'

    <#
    The range apply makes no undo entry. That is measured, not assumed: five
    Undos in a row leave all four Shapes filled, while one Undo reverts a
    single-Shape native apply, and one Undo reverts Office's own
    ShapeRange.Fill.UserPicture across all four.

    So Office does record undo for range fills - through PPCORE, above the
    receiver this path calls - and reaching the receiver directly skips it. That
    is the reason this cannot become what ApplyTexture does: a caller who fills
    32 Shapes and presses Ctrl+Z would get nothing back. It is asserted here as
    the behaviour it actually has, so that a future build changing it fails this
    test loudly rather than quietly gaining an undo entry nobody checked for.
    #>
    $presentation.Windows.Item(1).Activate()
    $app.StartNewUndoEntry()
    for ($u = 0; $u -lt 5; $u++) { $app.CommandBars.ExecuteMso('Undo') }
    Assert ((Count-Filled $slide $undoNames) -eq 4) `
        'the range apply makes NO undo entry - five Undos leave all four filled'

    # The control, in the same document and the same undo stack: a single-Shape
    # native apply is undone by one Undo. The difference is the range, not us.
    $control = $slide.Shapes.AddShape(1, 300, 380, 50, 50)
    $control.Name = 'UndoControl'
    $app.StartNewUndoEntry()
    $null = $engine.ApplyTexture($control, $handleA)
    Assert ($control.Fill.Type -eq 6) 'a single-Shape apply fills its Shape'
    $app.CommandBars.ExecuteMso('Undo')
    Assert ($control.Fill.Type -ne 6) 'and one Undo reverts it, so the undo stack is working'
    $control.Delete()

    # --- persistence ----------------------------------------------------------
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopenedSlide = $presentation.Slides.Item(1)
    $survived = ($names + $undoNames | ForEach-Object { $reopenedSlide.Shapes.Item($_).Fill.Type } |
        Where-Object { $_ -eq 6 }).Count
    Assert ($survived -eq 12) "every range-applied fill survives save and reopen ($survived of 12)"
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
    if (Test-Path $saved) { Remove-Item $saved -Force }
}

$results
if ($failures -gt 0) { throw "$failures check(s) failed" }
"all $($results.Count) checks passed"
