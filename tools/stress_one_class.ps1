<#
.SYNOPSIS
Stresses one Shape class through the native apply, in its own PowerPoint process.

.DESCRIPTION
A class that survives a single apply has not earned native support. This runs the
full sequence against one class and reports what moved:

  1000 repeated applies of one texture
  500  alternating applies of two textures
  undo and redo
  save, close and reopen
  delete the Shape, then keep applying to another
  close the presentation and check the reference count returns to the handle's own

Watched throughout: the cached image's reference count, private bytes, whether
the host is still alive, and whether a stale receiver is ever reached.

Classes the shipping allowlist already accepts go through the public
ApplyTexture. Candidates go through the research-only ApplyTextureUnrestricted,
which bypasses the allowlist and nothing else. Isolation is deliberate: if a
class takes the host down, the parent records it instead of losing the run.
#>
param(
    [Parameter(Mandatory = $true)][string]$Category,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Applies = 1000,
    [int]$Alternating = 500
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
function Get-PrivateMb {
    [math]::Round((Get-Process -Id $script:hostPid).PrivateMemorySize64 / 1MB, 1)
}

# The same factory table the compatibility probe uses. Kept in step with it by
# hand rather than shared, because a stress run must not be able to change what
# the matrix creates.
$factories = @{
    'AutoShape'   = { param($slide, $textSlide, $app) $slide.Shapes.AddShape(1, 20, 20, 120, 120) }
    'Freeform'    = { param($slide, $textSlide, $app)
        $b = $slide.Shapes.BuildFreeform(1, 200, 20)
        $b.AddNodes(1, 1, 320, 20); $b.AddNodes(1, 1, 320, 140); $b.AddNodes(1, 1, 200, 20)
        $b.ConvertToShape() }
    'TextBox'     = { param($slide, $textSlide, $app) $slide.Shapes.AddTextbox(1, 20, 200, 200, 60) }
    'Placeholder' = { param($slide, $textSlide, $app) $textSlide.Shapes.Placeholders.Item(1) }
    'Callout'     = { param($slide, $textSlide, $app) $slide.Shapes.AddCallout(2, 20, 300, 160, 60) }
    'Picture'     = { param($slide, $textSlide, $app)
        $slide.Shapes.AddPicture($texture, $false, $true, 540, 20, 80, 80) }
    'WordArt'     = { param($slide, $textSlide, $app)
        $slide.Shapes.AddTextEffect(1, 'BlipBridge', 'Arial', 24, $false, $false, 240, 300) }
    'Media'       = { param($slide, $textSlide, $app)
        $media = Join-Path $env:WINDIR 'Media\Windows Ding.wav'
        if (-not (Test-Path $media)) { throw 'no media file on this machine' }
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
# Classes the shipping allowlist accepts today; everything else needs the
# research path to be exercised at all.
$allowlisted = @('AutoShape', 'Freeform', 'WordArt', 'Group child',
                 'TextBox', 'Placeholder', 'Callout', 'Group', 'Picture', 'Media')

if (-not $factories.ContainsKey($Category)) {
    Set-Field 'result' 'UnknownCategory'
    Save-Result
    exit 0
}

Set-Field 'result' 'Crashed'
Set-Field 'stage' 'before start'
Save-Result

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$hostPid = $engine.GetHostProcessId()
$presentation = $app.Presentations.Add($msoTrue)
$useUnrestricted = -not ($allowlisted -contains $Category)
Set-Field 'path' $(if ($useUnrestricted) { 'unrestricted' } else { 'public' })

function Invoke-Apply($shape, $handle) {
    if ($script:useUnrestricted) { $null = $script:engine.ApplyTextureUnrestricted($shape, $handle) }
    else { $null = $script:engine.ApplyTexture($shape, $handle) }
}

try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $textSlide = $presentation.Slides.Add(2, $ppLayoutText)
    $shape = & $factories[$Category] $slide $textSlide $app

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    [byte[]]$bytesB = [IO.File]::ReadAllBytes($textureB)
    $handle = $engine.LoadTexture($bytes)
    $handleB = $engine.LoadTexture($bytesB)

    Set-Field 'stage' 'repeated applies'
    $privateStart = Get-PrivateMb
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $Applies; $i++) { Invoke-Apply $shape $handle }
    $watch.Stop()
    Set-Field 'applies' $Applies
    Set-Field 'msPerApply' ([math]::Round($watch.Elapsed.TotalMilliseconds / $Applies, 4))
    $report = $engine.InspectTexture($handle)
    Set-Field 'refsAfterApplies' (Get-Field $report 'cachedCount')
    Set-Field 'creations' (Get-Field $report 'creations')
    Set-Field 'privateAfterApplies' (Get-PrivateMb)

    Set-Field 'stage' 'alternating'
    for ($i = 0; $i -lt $Alternating; $i++) {
        if ($i % 2 -eq 0) { Invoke-Apply $shape $handle } else { Invoke-Apply $shape $handleB }
    }
    Set-Field 'refsAfterAlternating' (Get-Field ($engine.InspectTexture($handle)) 'cachedCount')

    Set-Field 'stage' 'undo/redo'
    $undo = 'no'
    try {
        $app.StartNewUndoEntry()
        Invoke-Apply $shape $handle
        $app.CommandBars.ExecuteMso('Undo')
        $app.CommandBars.ExecuteMso('Redo')
        $undo = 'yes'
    } catch { $undo = 'no' }
    Set-Field 'undoRedo' $undo

    Set-Field 'stage' 'save/reopen'
    $saved = Join-Path $env:TEMP ('bbstress_' + [guid]::NewGuid().ToString('N') + '.pptx')
    $reopen = 'failed'
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
                if ($top.Type -eq 6) { foreach ($member in $top.GroupItems) { $flat.Add($member) } }
            }
        }
        $filled = @($flat | Where-Object { try { $_.Fill.Type -eq 6 } catch { $false } }).Count
        $reopen = "filled $filled"
        $reopened.Saved = $msoTrue; $reopened.Close()
        $presentation = $anchor
        $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
        $textSlide = $presentation.Slides.Add(2, $ppLayoutText)
    } catch {
        $reopen = 'failed: ' + $_.Exception.Message
    } finally {
        Remove-Item $saved -ErrorAction SilentlyContinue
    }
    Set-Field 'reopen' $reopen

    Set-Field 'stage' 'delete and reapply'
    $deleteOk = 'no'
    try {
        $victim = & $factories[$Category] $slide $textSlide $app
        Invoke-Apply $victim $handle
        # A group child cannot be deleted directly; delete its group instead.
        if ($Category -eq 'Group child') { $victim.ParentGroup.Delete() } else { $victim.Delete() }
        $survivor = & $factories[$Category] $slide $textSlide $app
        Invoke-Apply $survivor $handle
        $deleteOk = 'yes'
    } catch { $deleteOk = 'no: ' + $_.Exception.Message }
    Set-Field 'deleteThenApply' $deleteOk

    Set-Field 'stage' 'close'
    try { $presentation.Saved = $msoTrue; $presentation.Close(); $presentation = $null } catch { }
    Set-Field 'refsAfterClose' (Get-Field ($engine.InspectTexture($handle)) 'cachedCount')

    $engine.ReleaseTexture($handleB)
    $engine.ReleaseTexture($handle)
    Set-Field 'handlesAtEnd' $engine.GetTextureCount()
    Set-Field 'privateStart' $privateStart
    Set-Field 'privateEnd' (Get-PrivateMb)

    # The host has to be usable afterwards, by its own API, not only by ours.
    $check = 'no'
    try {
        $fresh = $app.Presentations.Add($msoTrue)
        $s = $fresh.Slides.Add(1, $ppLayoutBlank)
        $probe = $s.Shapes.AddShape(1, 10, 10, 40, 40)
        $probe.Fill.UserPicture($texture)
        $check = 'yes'
        $fresh.Saved = $msoTrue; $fresh.Close()
    } catch { $check = 'no: ' + $_.Exception.Message }
    Set-Field 'hostHealthy' $check

    Set-Field 'stage' 'done'
    Set-Field 'result' 'Survived'
    Save-Result
} finally {
    if ($presentation) { try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { } }
    try { $app.Quit() } catch { }
}
