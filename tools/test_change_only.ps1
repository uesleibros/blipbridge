<#
.SYNOPSIS
Finds out what Office's change-recording step is buying, by leaving it out.

.DESCRIPTION
Taking the private apply apart showed where its cost is: performing the change
costs about 5 microseconds, and recording that change against the document costs
about 144 - more than nine tenths of the whole operation. A fill applied without
the recording costs 0.013 ms instead of 0.166.

That is only interesting if the document survives it, so this asks, one property
at a time, what the recording was for:

  fill        does the change reach Fill.Type and Fill.TextureType at all
  render      does Shape.Export produce the new image, or the old one
  modified    is the presentation marked as changed - because if it is not, a
              user can close it and lose work with no prompt
  persistence does the fill survive save and reopen
  undo        is there an undo entry, and what does Undo do if there is not
  stability   do a thousand of them in a row leave PowerPoint standing

Nothing here is reachable from the C ABI. This is a measurement of a code path,
not a proposal to ship one, and the results are what decide which it becomes.
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

$findings = New-Object System.Collections.Generic.List[string]
function Note([string]$what) { $script:findings.Add("  $what") }

$temp = [IO.Path]::GetTempPath()
$saved = Join-Path $temp ("bb_change_only_" + [Guid]::NewGuid().ToString('N') + '.pptx')
$shotBefore = Join-Path $temp ("bb_before_" + [Guid]::NewGuid().ToString('N') + '.png')
$shotAfter = Join-Path $temp ("bb_after_" + [Guid]::NewGuid().ToString('N') + '.png')

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

    # --- does the change reach the document model at all? ---------------------
    $shape = $slide.Shapes.AddShape(1, 20, 20, 160, 160)
    $shape.Name = 'BbChangeOnly'
    $shape.Export($shotBefore, 2)
    $sizeBefore = (Get-Item $shotBefore).Length

    # A real apply first, so "modified" below is about the second call and not
    # about having created the Shape.
    $null = $engine.ApplyTexture($shape, $handleA)
    $presentation.Saved = $msoTrue

    $null = $engine.ApplyChangeOnly($shape, $handleB)
    Note "Fill.Type after a change-only apply: $($shape.Fill.Type) (6 is a picture fill)"
    Note "Fill.TextureType: $($shape.Fill.TextureType)"

    $shape.Export($shotAfter, 2)
    $sizeAfter = (Get-Item $shotAfter).Length
    Note "rendered PNG: $sizeBefore bytes before any fill, $sizeAfter after - $(if ($sizeAfter -ne $sizeBefore) { 'the render changed' } else { 'the render did NOT change' })"

    # --- the modified flag ----------------------------------------------------
    Note "Presentation.Saved after a change-only apply: $($presentation.Saved) ($msoTrue means PowerPoint thinks nothing changed)"

    # --- undo -----------------------------------------------------------------
    # Whatever Undo does here, it must not crash and must not half-revert.
    $presentation.Windows.Item(1).Activate()
    $undoSurvived = $true
    try { $app.CommandBars.ExecuteMso('Undo') } catch { $undoSurvived = $false }
    Note "Undo after a change-only apply: $(if ($undoSurvived) { 'accepted' } else { 'refused' }), Fill.Type now $($shape.Fill.Type)"

    # --- persistence ----------------------------------------------------------
    $null = $engine.ApplyChangeOnly($shape, $handleA)
    $typeBeforeSave = $shape.Fill.Type
    $presentation.SaveAs($saved, $ppSaveAsDefault)
    $presentation.Close()
    $presentation = $app.Presentations.Open($saved, $msoFalse, $msoFalse, $msoTrue)
    $reopened = $presentation.Slides.Item(1).Shapes.Item('BbChangeOnly')
    Note "after save and reopen: Fill.Type $typeBeforeSave -> $($reopened.Fill.Type), TextureType $($reopened.Fill.TextureType)"

    # --- stability ------------------------------------------------------------
    $stress = $presentation.Slides.Item(1).Shapes.AddShape(1, 220, 20, 160, 160)
    $rounds = 1000
    for ($i = 0; $i -lt $rounds; $i++) {
        $null = $engine.ApplyChangeOnly($stress, $(if ($i % 2) { $handleA } else { $handleB }))
    }
    Note "$rounds change-only applies alternating two images: PowerPoint still answering, Fill.Type $($stress.Fill.Type)"

    # And an ordinary apply afterwards, to show the Shape is not left in a state
    # the supported path cannot deal with.
    $null = $engine.ApplyTexture($stress, $handleA)
    Note "an ordinary ApplyTexture afterwards: Fill.Type $($stress.Fill.Type)"
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
    foreach ($file in @($saved, $shotBefore, $shotAfter)) {
        if (Test-Path $file) { Remove-Item $file -Force }
    }
}

'change-only apply, what the recording was buying:'
$findings
