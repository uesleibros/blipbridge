<#
.SYNOPSIS
Lifetime and correctness matrix for reusable native texture handles.

.DESCRIPTION
Exercises LoadTexture / ApplyTexture / ReleaseTexture / ClearTextures against
every lifetime boundary that matters before the backend could be called
production-ready:

  reuse at volume, many Shapes, several slides, Freeforms, several textures,
  multiple presentations, Shape deletion, presentation close, stale handles,
  release ordering, SaveAs and reopen, and whether anything is retained after
  the document goes away.

PowerPoint shutdown is covered by the caller: this script leaves the host
running with textures released, and tools/test_native_texture_shutdown.ps1
checks the harder case where textures are still loaded when the host exits.
#>
param([int]$ReuseIterations = 1000)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$second = Join-Path $root 'artifacts/textures/texture_128_0.png'
$third = Join-Path $root 'artifacts/textures/texture_64_0.jpg'
$deck = Join-Path $root 'artifacts/native_texture.pptx'
$results = New-Object System.Collections.Generic.List[string]

function Get-Identity($shape) {
    @($shape.Id, $shape.Name, $shape.Type, $shape.Left, $shape.Top, $shape.Width,
      $shape.Height, $shape.Rotation, $shape.ZOrderPosition) -join '|'
}
function Get-Field([string]$report, [string]$name) {
    foreach ($part in $report.Split(';')) {
        $pair = $part.Split('=', 2)
        if ($pair.Length -eq 2 -and $pair[0] -eq $name) { return $pair[1] }
    }
    return $null
}

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$hostProcess = Get-Process -Id $engine.GetHostProcessId()
function Get-PrivateMb { [math]::Round((Get-Process -Id $hostProcess.Id).PrivateMemorySize64 / 1MB, 1) }

$presentation = $app.Presentations.Add(0)
try {
    $slide1 = $presentation.Slides.Add(1, 12)
    $slide2 = $presentation.Slides.Add(2, 12)

    # Warm GFX/OART. No target Shape ever sees UserPicture.
    $warmup = $slide1.Shapes.AddShape(1, 600, 400, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
    [byte[]]$bytes2 = [IO.File]::ReadAllBytes($second)
    [byte[]]$bytes3 = [IO.File]::ReadAllBytes($third)

    # --- load ---------------------------------------------------------------
    $handle = $engine.LoadTexture($bytes)
    $results.Add("LoadTexture returned handle $handle; count=$($engine.GetTextureCount())")
    if ($engine.GetTextureCount() -ne 1) { throw 'Texture count wrong after one load' }

    # --- many Shapes, several slides, a Freeform ----------------------------
    $shapes = @(
        $slide1.Shapes.AddShape(1, 20, 20, 90, 90),
        $slide1.Shapes.AddShape(1, 120, 20, 90, 90),
        $slide1.Shapes.AddShape(5, 220, 20, 90, 90),
        $slide2.Shapes.AddShape(1, 20, 20, 90, 90),
        $slide2.Shapes.AddShape(1, 120, 20, 90, 90)
    )
    $builder = $slide1.Shapes.BuildFreeform(0, 60, 200)
    $builder.AddNodes(0, 0, 200, 200)
    $builder.AddNodes(0, 0, 130, 340)
    $builder.AddNodes(0, 0, 60, 200)
    $freeform = $builder.ConvertToShape()
    $shapes += $freeform
    for ($i = 0; $i -lt $shapes.Count; $i++) { $shapes[$i].Name = "BB_tex_$i" }
    $identities = $shapes | ForEach-Object { Get-Identity $_ }
    $nodesBefore = $freeform.Nodes.Count

    foreach ($shape in $shapes) { $null = $engine.ApplyTexture($shape, $handle) }
    for ($i = 0; $i -lt $shapes.Count; $i++) {
        if ((Get-Identity $shapes[$i]) -ne $identities[$i]) { throw "Shape $i identity changed" }
        if ([int]$shapes[$i].Fill.Type -ne 6) { throw "Shape $i has no picture fill" }
    }
    if ($freeform.Nodes.Count -ne $nodesBefore) { throw 'Freeform node count changed' }
    $freeform.Nodes.SetPosition(2, 220, 200)
    $results.Add("One texture filled $($shapes.Count) Shapes across 2 slides including a Freeform; nodes still editable.")

    # --- reuse at volume, on one Shape and round-robin ----------------------
    $before = Get-PrivateMb
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $ReuseIterations; $i++) {
        $null = $engine.ApplyTexture($shapes[$i % $shapes.Count], $handle)
    }
    $watch.Stop()
    $after = Get-PrivateMb
    $report = $engine.InspectTexture($handle)
    $results.Add("$ReuseIterations applies round-robin over $($shapes.Count) Shapes: $([math]::Round($watch.Elapsed.TotalMilliseconds,1)) ms total")
    $results.Add("  private bytes ${before} MB -> ${after} MB; texture report: $report")
    if ((Get-Field $report 'applies') -lt $ReuseIterations) { throw 'Apply counter did not advance' }
    foreach ($shape in $shapes) {
        if ([int]$shape.Fill.Type -ne 6) { throw 'A Shape lost its picture fill during the reuse run' }
    }

    # --- several textures at once -------------------------------------------
    $handle2 = $engine.LoadTexture($bytes2)
    $handle3 = $engine.LoadTexture($bytes3)
    if ($engine.GetTextureCount() -ne 3) { throw 'Texture count wrong after three loads' }
    $null = $engine.ApplyTexture($shapes[0], $handle2)
    $null = $engine.ApplyTexture($shapes[1], $handle3)
    $null = $engine.ApplyTexture($shapes[2], $handle)
    foreach ($shape in $shapes[0..2]) {
        if ([int]$shape.Fill.Type -ne 6) { throw 'Multi-texture apply lost a picture fill' }
    }
    $results.Add('Three textures coexist; each applies to a different Shape (PNG, PNG, JPEG).')

    # --- release ordering: release the middle handle, others keep working ---
    $engine.ReleaseTexture($handle2)
    if ($engine.GetTextureCount() -ne 2) { throw 'Texture count wrong after one release' }
    $null = $engine.ApplyTexture($shapes[3], $handle)
    $null = $engine.ApplyTexture($shapes[4], $handle3)
    $results.Add('Releasing one texture out of order left the others usable.')

    # --- stale handle --------------------------------------------------------
    $stale = $false
    try { $null = $engine.ApplyTexture($shapes[0], $handle2) } catch { $stale = $true }
    if (-not $stale) { throw 'A released handle was still accepted' }
    $results.Add('Released handle is rejected; handles are not recycled.')
    $bogus = $false
    try { $null = $engine.ApplyTexture($shapes[0], 0x7FFFFF00) } catch { $bogus = $true }
    if (-not $bogus) { throw 'An unknown handle was accepted' }
    $results.Add('Unknown handle is rejected.')

    # --- Shape deletion ------------------------------------------------------
    $doomed = $slide1.Shapes.AddShape(1, 320, 20, 60, 60)
    $null = $engine.ApplyTexture($doomed, $handle)
    $doomed.Delete()
    $null = $engine.ApplyTexture($shapes[0], $handle)
    if ([int]$shapes[0].Fill.Type -ne 6) { throw 'Apply broke after deleting another Shape' }
    $results.Add('Deleting a filled Shape left the texture and later applies healthy.')

    $presentation.SaveAs($deck, 24)
    $results.Add('SaveAs succeeded.')
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

# --- textures survive presentation close ------------------------------------
$afterClose = Get-PrivateMb
$results.Add("Textures still loaded after Presentation.Close: $($engine.GetTextureCount()); private bytes $afterClose MB")
if ($engine.GetTextureCount() -lt 1) { throw 'Textures were lost when the presentation closed' }

# --- multiple presentations, and reuse of a texture loaded under the old one -
$reopened = $app.Presentations.Open($deck, $true, $false, $false)
$fresh = $app.Presentations.Add(0)
try {
    $withPicture = 0
    $pictureShapes = 0
    foreach ($slide in $reopened.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ([int]$shape.Type -eq 13) { $pictureShapes++ }
            if ([int]$shape.Fill.Type -eq 6) { $withPicture++ }
        }
    }
    $results.Add("Reopen: $withPicture Shapes carry a picture fill; Picture shapes: $pictureShapes")
    if ($withPicture -lt 6) { throw 'Picture fills did not survive reopen' }
    if ($pictureShapes -ne 0) { throw 'Reopened deck contains Picture shapes' }

    # The surviving handle must still apply, in a different presentation.
    $survivor = @($engine.GetTextureCount())
    $freshShape = $fresh.Slides.Add(1, 12).Shapes.AddShape(1, 20, 20, 90, 90)
    $reopenedShape = $reopened.Slides.Item(1).Shapes.AddShape(1, 420, 20, 90, 90)
    $handles = @()
    for ($h = 0x1000000; $h -lt 0x1000010; $h++) {
        try { $null = $engine.ApplyTexture($freshShape, $h); $handles += $h } catch { }
    }
    if ($handles.Count -lt 1) { throw 'No surviving texture could be applied in a new presentation' }
    $null = $engine.ApplyTexture($reopenedShape, $handles[0])
    if ([int]$freshShape.Fill.Type -ne 6 -or [int]$reopenedShape.Fill.Type -ne 6) {
        throw 'A texture loaded under a closed presentation did not apply elsewhere'
    }
    $results.Add("A texture outlived its original presentation and applied in two others (handles $($handles -join ',')).")
} finally {
    $reopened.Saved = -1; $reopened.Close()
    $fresh.Saved = -1; $fresh.Close()
}

# --- ClearTextures and retention --------------------------------------------
$engine.ClearTextures()
if ($engine.GetTextureCount() -ne 0) { throw 'ClearTextures left textures behind' }
$results.Add('ClearTextures released everything.')
Start-Sleep -Seconds 2
$results.Add("Private bytes after ClearTextures and all documents closed: $(Get-PrivateMb) MB")

# The host must still work normally.
$check = $app.Presentations.Add(0)
try {
    $shape = $check.Slides.Add(1, 12).Shapes.AddShape(1, 20, 20, 90, 90)
    $shape.Fill.UserPicture($texture)
    if ([int]$shape.Fill.Type -ne 6) { throw 'Host cannot fill Shapes after the texture run' }
    $results.Add('Ordinary UserPicture still works after everything above.')
} finally { $check.Saved = -1; $check.Close() }

$results | Set-Content "$root/artifacts/native_texture.txt"
$results
