<#
.SYNOPSIS
Runs one semantic-guard regression case in its own PowerPoint process.

.DESCRIPTION
Called by tools/test_semantic_guards.ps1. Each case is a Shape class plus the
outcome the safety model promises for it, and the crucial assertion is not "the
host survived" but **"the private OART apply was never entered"**.

That is provable rather than inferred: the backend counts entries into
ApplyCachedImage, and the probe reports the count. A refusal that happens before
the dangerous call leaves the count unchanged. A refusal that happens after would
show up here even if PowerPoint happened to survive it - which is exactly the
failure mode a "did it crash?" test cannot see.

Cases:

  Connector  Unsupported, refused before the private apply, host survives
  Line       Unsupported, refused before the private apply, host survives
  WordArt    NativeSupported: applies, stays WordArt, keeps its text and
             geometry, persists through save/reopen, undoes and redoes

Writes one `key=value;` line to -Out. Isolated because a regression here is a
crash, and a crash must cost one row rather than the whole suite.
#>
param(

    [Parameter(Mandatory = $true)][ValidateSet('Connector', 'Line', 'WordArt')][string]$Case,
    [Parameter(Mandatory = $true)][string]$Out
)
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
$ppLayoutBlank = 12
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'

$fields = [ordered]@{ case = $Case }
function Set-Field([string]$name, $value) {
    $script:fields[$name] = ("$value" -replace '[;=\r\n]', ' ')
}
function Save-Result {
    $text = (($script:fields.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ';')
    Set-Content -Path $Out -Value $text
}
function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}
$failures = 0
function Assert([bool]$condition, [string]$what) {
    if (-not $condition) {
        $script:failures++
        $script:fields['failed'] = "$($script:fields['failed']) | $what".Trim(' |')
    }
}

# Written first, so a host that dies leaves a row saying so.
Set-Field 'result' 'Crashed'
Save-Result


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    $handle = $engine.LoadTexture($bytes)
    Set-Field 'handle' $handle

    switch ($Case) {
        'Connector' { $shape = $slide.Shapes.AddConnector(1, 60, 60, 260, 200) }
        'Line'      { $shape = $slide.Shapes.AddLine(60, 240, 260, 320) }
        'WordArt'   {
            $shape = $slide.Shapes.AddTextEffect(1, 'BlipBridge', 'Arial', 28, $false, $false, 60, 60)
        }
    }
    Set-Field 'shapeType' $shape.Type

    # The verdict and the entry count, before anything is attempted.
    $before = $engine.ProbeShapePolicy($shape)
    $eligibility = Get-Field $before 'eligibility'
    $entriesBefore = [int](Get-Field $before 'applyEntries')
    Set-Field 'eligibility' $eligibility
    Set-Field 'connector' (Get-Field $before 'connector')
    Set-Field 'entriesBefore' $entriesBefore

    if ($Case -eq 'WordArt') {
        # --- the opposite problem: looks ordinary, must actually work ----------
        Assert ($eligibility -eq 'NativeSupported') 'WordArt is NativeSupported'
        $textBefore = ''
        try { $textBefore = $shape.TextEffect.Text } catch { $textBefore = '<unreadable>' }
        $geometry = '{0},{1},{2},{3}' -f [math]::Round($shape.Left, 3), [math]::Round($shape.Top, 3),
                                          [math]::Round($shape.Width, 3), [math]::Round($shape.Height, 3)
        $fillBefore = $shape.Fill.Type
        Set-Field 'fillBefore' $fillBefore

        $null = $engine.ApplyTexture($shape, $handle)
        $entriesAfter = [int](Get-Field ($engine.ProbeShapePolicy($shape)) 'applyEntries')
        Set-Field 'entriesAfter' $entriesAfter
        Assert ($entriesAfter -gt $entriesBefore) 'the native apply really ran'

        Set-Field 'fillAfter' $shape.Fill.Type
        Assert ($shape.Fill.Type -eq 6) 'WordArt gets a picture fill'
        Assert ($shape.Type -eq 15 -or $shape.Type -eq 1) 'WordArt keeps its Shape.Type'
        $textAfter = ''
        try { $textAfter = $shape.TextEffect.Text } catch { $textAfter = '<unreadable>' }
        Set-Field 'text' $textAfter
        Assert ($textAfter -eq $textBefore) 'the WordArt text is unchanged'
        $geometryAfter = '{0},{1},{2},{3}' -f [math]::Round($shape.Left, 3), [math]::Round($shape.Top, 3),
                                               [math]::Round($shape.Width, 3), [math]::Round($shape.Height, 3)
        Assert ($geometryAfter -eq $geometry) 'the WordArt geometry is unchanged'
        Assert (@($slide.Shapes | Where-Object { $_.Type -eq 13 }).Count -eq 0) `
            'no Picture Shape was substituted'

        # undo, redo
        $app.StartNewUndoEntry()
        $null = $engine.ApplyTexture($shape, $handle)
        $app.CommandBars.ExecuteMso('Undo')
        $undone = $shape.Fill.Type
        $app.CommandBars.ExecuteMso('Redo')
        $redone = $shape.Fill.Type
        Set-Field 'undoRedo' "$undone/$redone"
        Assert ($redone -eq 6) 'redo restores the picture fill'

        # save, reopen
        $saved = Join-Path $env:TEMP ('bbword_' + [guid]::NewGuid().ToString('N') + '.pptx')
        try {
            $presentation.SaveAs($saved)
            $anchor = $app.Presentations.Add($msoTrue)
            $presentation.Close()
            $presentation = $null
            $reopened = $app.Presentations.Open($saved, $false, $false, $msoTrue)
            $survivor = $reopened.Slides.Item(1).Shapes.Item(1)
            Set-Field 'reopenFill' $survivor.Fill.Type
            Set-Field 'reopenType' $survivor.Type
            Assert ($survivor.Fill.Type -eq 6) 'the picture fill survives save and reopen'
            $reopenText = ''
            try { $reopenText = $survivor.TextEffect.Text } catch { $reopenText = '<unreadable>' }
            Assert ($reopenText -eq $textBefore) 'and so does the WordArt text'
            $reopened.Saved = $msoTrue; $reopened.Close()
            $presentation = $anchor
        } finally {
            Remove-Item $saved -ErrorAction SilentlyContinue
        }
    } else {
        # --- the dangerous classes: refused before anything internal happens ---
        Assert ($eligibility -eq 'Unsupported') "$Case is classified Unsupported"
        Assert ((Get-Field $before 'connector') -eq '1') "$Case reports Shape.Connector"

        $refusal = ''
        try { $null = $engine.ApplyTexture($shape, $handle) }
        catch { $refusal = $_.Exception.Message }
        Set-Field 'refusal' $refusal
        Assert ($refusal -ne '') "$Case is refused rather than applied"

        # The assertion this whole file exists for.
        $entriesAfter = [int](Get-Field ($engine.ProbeShapePolicy($shape)) 'applyEntries')
        Set-Field 'entriesAfter' $entriesAfter
        Assert ($entriesAfter -eq $entriesBefore) `
            'the private OART apply was never entered'

        # And the same through the one-call path, which must not fall back either.
        $pictureRefusal = ''
        try { $null = $engine.ApplyPicture($shape, $texture) }
        catch { $pictureRefusal = $_.Exception.Message }
        Set-Field 'pictureRefusal' $pictureRefusal
        Assert ($pictureRefusal -ne '') "UserPicture2 refuses $Case too"
        $entriesFinal = [int](Get-Field ($engine.ProbeShapePolicy($shape)) 'applyEntries')
        Assert ($entriesFinal -eq $entriesBefore) `
            'and UserPicture2 did not enter the private apply either'

        Set-Field 'fillAfter' $shape.Fill.Type
        Assert ($shape.Type -ne 13) 'the Shape was not replaced'
    }

    # The host has to be usable by its own API afterwards, not just by ours.
    $healthy = 'no'
    try {
        if (-not $presentation) { $presentation = $app.Presentations.Add($msoTrue) }
        $check = $presentation.Slides.Add(1, $ppLayoutBlank).Shapes.AddShape(1, 10, 10, 40, 40)
        $check.Fill.UserPicture($texture)
        if ($check.Fill.Type -eq 6) { $healthy = 'yes' }
    } catch { $healthy = 'no: ' + $_.Exception.Message }
    Set-Field 'hostHealthy' $healthy
    Assert ($healthy -eq 'yes') 'PowerPoint is healthy afterwards'

    $verdict = 'Failed'
    if ($failures -eq 0) { $verdict = 'Passed' }
    Set-Field 'result' $verdict
    Save-Result
} finally {
    if ($presentation) { try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { } }
    try { $app.Quit() } catch { }
}
