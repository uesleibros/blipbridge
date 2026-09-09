<#
.SYNOPSIS
Tests exactly one Shape category, in its own PowerPoint process.

.DESCRIPTION
Called by tools/test_shape_compatibility.ps1, once per category, as a separate
process. That isolation is the point: a class whose native apply terminates
PowerPoint - and one has already been found - would otherwise destroy every
result gathered after it. Here, the parent simply sees a child that never wrote
its row, and records `Crashed`.

Writes one `key=value;` line to -Out and exits. Never throws for a Shape class
that is merely unsupported; that is a result, not an error.

.PARAMETER Category
The category name to test. See $factories below.

.PARAMETER Out
File to write the single result line to.
#>
param(
    [Parameter(Mandatory = $true)][string]$Category,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

$msoTrue = -1
$ppLayoutBlank = 12
$ppLayoutText = 2
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'
$textureB = Join-Path $root 'artifacts/textures/texture_64_1.png'

$fields = [ordered]@{ category = $Category }
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

<#
 Each factory receives the blank slide, the text-layout slide and the app, and
 returns the Shape to test. A factory may legitimately fail on a machine without
 the relevant Office component; that becomes NotApplicable, not a failure.
#>
$factories = @{
    'AutoShape'   = { param($slide, $textSlide, $app) $slide.Shapes.AddShape(1, 20, 20, 120, 120) }
    'Freeform'    = { param($slide, $textSlide, $app)
        $b = $slide.Shapes.BuildFreeform(1, 200, 20)
        $b.AddNodes(1, 1, 320, 20); $b.AddNodes(1, 1, 320, 140); $b.AddNodes(1, 1, 200, 20)
        $b.ConvertToShape() }
    'TextBox'     = { param($slide, $textSlide, $app) $slide.Shapes.AddTextbox(1, 20, 200, 200, 60) }
    'Placeholder' = { param($slide, $textSlide, $app) $textSlide.Shapes.Placeholders.Item(1) }
    'Connector'   = { param($slide, $textSlide, $app) $slide.Shapes.AddConnector(1, 400, 20, 500, 120) }
    'Line'        = { param($slide, $textSlide, $app) $slide.Shapes.AddLine(400, 200, 500, 260) }
    'Callout'     = { param($slide, $textSlide, $app) $slide.Shapes.AddCallout(2, 20, 300, 160, 60) }
    'Picture'     = { param($slide, $textSlide, $app)
        $slide.Shapes.AddPicture($texture, $false, $true, 540, 20, 80, 80) }
    'WordArt'     = { param($slide, $textSlide, $app)
        $slide.Shapes.AddTextEffect(1, 'BlipBridge', 'Arial', 24, $false, $false, 240, 300) }
    'Table'       = { param($slide, $textSlide, $app) $slide.Shapes.AddTable(2, 2, 20, 380, 200, 80) }
    'Chart'       = { param($slide, $textSlide, $app) $slide.Shapes.AddChart2(-1, 51, 300, 380, 200, 150) }
    'SmartArt'    = { param($slide, $textSlide, $app)
        $slide.Shapes.AddSmartArt($app.SmartArtLayouts.Item(1), 540, 380, 180, 120) }
    'OLE'         = { param($slide, $textSlide, $app)
        $slide.Shapes.AddOLEObject(20, 480, 160, 90, 'Package') }
    'Media'       = { param($slide, $textSlide, $app)
        # A real media file is required; Windows ships one under Media.
        $media = Join-Path $env:WINDIR 'Media\Windows Ding.wav'
        if (-not (Test-Path $media)) { throw "no media file on this machine" }
        $slide.Shapes.AddMediaObject2($media, $false, $true, 300, 200, 80, 80) }
    'Group'       = { param($slide, $textSlide, $app)
        $a = $slide.Shapes.AddShape(1, 640, 200, 60, 60)
        $b = $slide.Shapes.AddShape(1, 640, 280, 60, 60)
        $slide.Shapes.Range(@($a.Name, $b.Name)).Group() }
    'Group child' = { param($slide, $textSlide, $app)
        $a = $slide.Shapes.AddShape(1, 640, 200, 60, 60)
        $b = $slide.Shapes.AddShape(1, 640, 280, 60, 60)
        $slide.Shapes.Range(@($a.Name, $b.Name)).Group().GroupItems.Item(1) }
}

if (-not $factories.ContainsKey($Category)) {
    Set-Field 'class' 'Unsupported'
    Set-Field 'step' 'UnknownCategory'
    Save-Result
    exit 0
}

# Written before anything risky, so a crash still leaves a row saying where.
Set-Field 'class' 'Crashed'
Set-Field 'step' 'HostDiedDuringTest'
Save-Result

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $textSlide = $presentation.Slides.Add(2, $ppLayoutText)

    $shape = $null
    try {
        $shape = & $factories[$Category] $slide $textSlide $app
    } catch {
        Set-Field 'class' 'NotApplicable'
        Set-Field 'step' 'CreateFailed'
        Set-Field 'detail' $_.Exception.Message
        Save-Result
        exit 0
    }

    Set-Field 'shapeType' $shape.Type
    $connector = 'throws'
    try { $connector = $shape.Connector } catch { }
    Set-Field 'connector' $connector

    $report = $engine.ProbeShapeCompatibility($shape)
    $step = Get-Field $report 'step'
    Set-Field 'step' $step
    Set-Field 'agrees' (Get-Field $report 'agrees')
    $inner = Get-Field $report 'innerVtableRva'
    if ($inner) { Set-Field 'inner' $inner }

    if ((Get-Field $report 'agrees') -ne '1') {
        Set-Field 'class' 'Unsupported'
        Set-Field 'step' 'ProbeInconsistent'
        Set-Field 'detail' 'classifier disagreed with ResolveFillTarget'
        Save-Result
        exit 0
    }

    # The fallback question is asked for every class, supported or not: the
    # UserPicture2 dispatcher needs to know whether falling back is even an
    # option. It is asked on a *separate* Shape so a successful fallback cannot
    # contaminate the native measurement.
    $fallbackShape = $null
    try { $fallbackShape = & $factories[$Category] $slide $textSlide $app } catch { }
    $fallback = 'no'
    if ($fallbackShape) {
        try { $fallbackShape.Fill.UserPicture($texture); $fallback = 'yes' }
        catch { $fallback = 'no' }
    }
    Set-Field 'fallback' $fallback

    if ($step -ne 'Complete') {
        $class = 'Unsupported'
        if ($fallback -eq 'yes') { $class = 'FallbackSupported' }
        Set-Field 'class' $class
        Set-Field 'detail' 'no validated receiver chain'
        Save-Result
        exit 0
    }

    # Structure is intact. Whether that means it is safe is what the apply finds
    # out - and the row above already says Crashed if it is not.
    $before = @{
        Name = $shape.Name; Type = $shape.Type; Id = $shape.Id
        Left = [math]::Round($shape.Left, 3); Top = [math]::Round($shape.Top, 3)
        Width = [math]::Round($shape.Width, 3); Height = [math]::Round($shape.Height, 3)
        Rotation = $shape.Rotation
        Pictures = @($slide.Shapes | Where-Object { $_.Type -eq 13 }).Count
    }

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    [byte[]]$bytesB = [IO.File]::ReadAllBytes($textureB)
    $handle = $engine.LoadTexture($bytes)
    $handleB = $engine.LoadTexture($bytesB)

    $applyError = $null
    try { $null = $engine.ApplyTexture($shape, $handle) }
    catch { $applyError = $_.Exception.Message }

    <#
     The shipping allowlist refused it, but the validated chain is present. That
     is exactly the situation the Connector proved cannot be reasoned about - so
     it is tried, on a *separate* Shape of the same class, through the
     research-only path that bypasses the allowlist and nothing else.

     This process may not survive that. It is meant not to matter: the parent
     records `Crashed` for the class, which is the finding.
    #>
    if ($applyError) {
        Set-Field 'step' 'ApplyRefused'
        Set-Field 'detail' $applyError

        $unrestricted = 'notTried'
        if ($applyError -like '*no validated picture-fill path*') {
            $probeShape = $null
            try { $probeShape = & $factories[$Category] $slide $textSlide $app } catch { }
            if ($probeShape) {
                # Counted *after* the probe Shape exists. The Picture factory
                # adds a Picture, and baselining before it would blame the apply
                # for a Shape the test itself created.
                $picturesBefore = @($slide.Shapes | Where-Object { $_.Type -eq 13 }).Count
                $unrestricted = 'refused'
                try {
                    $null = $engine.ApplyTextureUnrestricted($probeShape, $handle)
                    $unrestricted = 'applied'
                } catch {
                    $unrestricted = 'refused: ' + $_.Exception.Message
                }
                if ($unrestricted -eq 'applied') {
                    # Survived the call. Now the same verification the allowlisted
                    # classes get, because surviving is not the same as correct.
                    $issues = New-Object System.Collections.Generic.List[string]
                    try {
                        if ($probeShape.Type -ne $shape.Type) { $issues.Add('type changed') }
                        if ($probeShape.Fill.Type -ne 6) {
                            $issues.Add("Fill.Type=$($probeShape.Fill.Type) not picture")
                        }
                    } catch { $issues.Add('shape unreadable after apply') }
                    if (@($slide.Shapes | Where-Object { $_.Type -eq 13 }).Count -ne $picturesBefore) {
                        $issues.Add('a Picture Shape appeared')
                    }
                    if ($issues.Count -gt 0) {
                        $unrestricted = 'applied but wrong: ' + ($issues -join ' / ')
                    } else {
                        $unrestricted = 'applied and verified'
                    }
                }
            }
        }
        Set-Field 'unrestricted' $unrestricted

        $class = 'Unsupported'
        if ($fallback -eq 'yes') { $class = 'FallbackSupported' }
        # A class the allowlist refuses but which applies and verifies cleanly is
        # a *candidate* for native support, not a decision. Promoting it happens
        # in the backend, deliberately, after the stress suite has run.
        if ($unrestricted -eq 'applied and verified') { $class = 'NativeCandidate' }
        Set-Field 'class' $class
        Save-Result
        exit 0
    }

    $problems = New-Object System.Collections.Generic.List[string]
    if ($shape.Name -ne $before.Name) { $problems.Add('name changed') }
    if ($shape.Type -ne $before.Type) { $problems.Add("type $($before.Type)->$($shape.Type)") }
    if ($shape.Id -ne $before.Id) { $problems.Add('id changed') }
    foreach ($metric in 'Left','Top','Width','Height') {
        if ([math]::Round($shape.$metric, 3) -ne $before[$metric]) { $problems.Add("$metric changed") }
    }
    if ($shape.Rotation -ne $before.Rotation) { $problems.Add('rotation changed') }
    if (@($slide.Shapes | Where-Object { $_.Type -eq 13 }).Count -ne $before.Pictures) {
        $problems.Add('a Picture Shape appeared')
    }
    $fillType = $null
    try { $fillType = $shape.Fill.Type } catch { $problems.Add('Fill.Type unreadable') }
    if ($fillType -ne 6) { $problems.Add("Fill.Type=$fillType not picture") }
    Set-Field 'fillType' $fillType

    try { $null = $engine.ApplyTexture($shape, $handleB) } catch { $problems.Add('second apply failed') }

    $undo = 'no'
    try {
        $app.StartNewUndoEntry()
        $null = $engine.ApplyTexture($shape, $handle)
        $app.CommandBars.ExecuteMso('Undo')
        $undo = 'yes'
    } catch { $undo = 'no' }
    Set-Field 'undo' $undo

    # Save and reopen, in this child, so the row carries its own persistence
    # evidence rather than relying on a shared document at the end of a run.
    $saved = Join-Path $env:TEMP ('bb_' + [guid]::NewGuid().ToString('N') + '.pptx')
    $reopen = 'no'
    try {
        $presentation.SaveAs($saved)
        $anchor = $app.Presentations.Add($msoTrue)
        $presentation.Close()
        $presentation = $null
        $reopened = $app.Presentations.Open($saved, $false, $false, $msoTrue)
        $flat = New-Object System.Collections.Generic.List[object]
        foreach ($reopenedSlide in $reopened.Slides) {
            foreach ($top in $reopenedSlide.Shapes) {
                $flat.Add($top)
                if ($top.Type -eq 6) {
                    # A group child's fill is not visible from the group, so a
                    # group-child row would otherwise report zero and read as a
                    # persistence failure that did not happen. Placeholders sit
                    # on the second slide, which is why every slide is walked.
                    foreach ($member in $top.GroupItems) { $flat.Add($member) }
                }
            }
        }
        $filled = @($flat | Where-Object { try { $_.Fill.Type -eq 6 } catch { $false } }).Count
        $strays = @($flat | Where-Object { $_.Type -eq 13 }).Count
        $reopen = "filled=$filled,pictureShapes=$strays"
        $reopened.Saved = $msoTrue; $reopened.Close()
        $anchor.Saved = $msoTrue; $anchor.Close()
    } catch {
        $reopen = 'failed'
    } finally {
        Remove-Item $saved -ErrorAction SilentlyContinue
    }
    Set-Field 'reopen' $reopen

    if ($problems.Count -gt 0) {
        Set-Field 'class' 'Unsupported'
        Set-Field 'step' 'VerificationFailed'
        Set-Field 'detail' ($problems -join ' / ')
    } else {
        Set-Field 'class' 'NativeSupported'
        Set-Field 'detail' 'identity, geometry and type intact'
    }
    Save-Result
} finally {
    if ($presentation) { try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { } }
    try { $app.Quit() } catch { }
}
