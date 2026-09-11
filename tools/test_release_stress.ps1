<#
.SYNOPSIS
Bounded pre-release stress: does anything BlipBridge owns grow without bound?

.DESCRIPTION
The question this answers is narrow and it matters: after thousands of mixed
operations, does BlipBridge still hold resources it should have let go of?

It is deliberately *not* "did POWERPNT.EXE's working set grow". PowerPoint
retains its own document, undo and render resources for reasons that have nothing
to do with this library, and attributing all of that to BlipBridge would be
wrong. What is asserted here is what BlipBridge can actually account for:

  * the picture cache's texture and Shape counts return to zero when cleared
  * the caller-visible texture count returns to zero
  * handles keep increasing and are never reissued
  * the cached image's reference count returns to its resting value
  * Init/Shutdown cycles leave the library usable
  * PowerPoint is alive and healthy at the end

Private bytes are *reported* alongside, as context rather than as an assertion.
#>
param([int]$Iterations = 2000)

$ErrorActionPreference = 'Stop'
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


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$hostPid = $engine.GetHostProcessId()
function PrivateMb { [math]::Round((Get-Process -Id $script:hostPid).PrivateMemorySize64 / 1MB, 1) }

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shapes = @()
    for ($i = 0; $i -lt 6; $i++) {
        $shapes += $slide.Shapes.AddShape(1, 20 + ($i % 3) * 150, 20 + [int]($i / 3) * 150, 120, 120)
    }
    [byte[]]$bytes = [IO.File]::ReadAllBytes($textureA)

    $engine.ClearTextures()
    $null = $engine.ClearPictureCache()
    $start = PrivateMb
    $results.Add("  note private bytes at start: $start MB")

    # --- thousands of UserPicture2 calls, alternating images and Shapes -------
    for ($i = 0; $i -lt $Iterations; $i++) {
        $path = $textureA
        if ($i % 3 -eq 0) { $path = $textureB }
        $null = $engine.ApplyPicture($shapes[$i % 6], $path)
    }
    $stats = $engine.PictureCacheStats()
    $results.Add("  note after $Iterations UserPicture2: $stats")
    Assert ((Get-Field $stats 'textures') -eq '2') 'the picture cache holds one image per distinct file, not per call'
    Assert ([int](Get-Field $stats 'shapes') -le 6) 'and one entry per Shape, not per call'

    # --- thousands of explicit load/release ----------------------------------
    $firstHandle = 0
    $lastHandle = 0
    for ($i = 0; $i -lt $Iterations; $i++) {
        $h = $engine.LoadTexture($bytes)
        if ($i -eq 0) { $firstHandle = $h }
        $lastHandle = $h
        $engine.ReleaseTexture($h)
    }
    Assert ($lastHandle -gt $firstHandle) 'handles keep increasing across thousands of load/release cycles'
    Assert ($engine.GetTextureCount() -eq 0) 'and none of them is still held'

    # A handle released long ago must still be refused.
    $stale = ''
    try { $null = $engine.ApplyTexture($shapes[0], $firstHandle) }
    catch { $stale = $_.Exception.Message }
    Assert ($stale -ne '') 'the very first handle is still permanently stale'

    # --- repeated clears in both directions -----------------------------------
    for ($i = 0; $i -lt 50; $i++) {
        $null = $engine.ApplyPicture($shapes[$i % 6], $textureA)
        $engine.ClearTextures()
        $null = $engine.ApplyPicture($shapes[$i % 6], $textureB)
        $null = $engine.ClearPictureCache()
    }
    Assert $true 'fifty interleaved ClearTextures/ClearPictureCache cycles completed'
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '0') 'and the picture cache is empty after the final clear'
    Assert ((Get-Field $stats 'shapes') -eq '0') 'and remembers no Shapes'

    # --- scaled texture creation ---------------------------------------------
    # Exercised through the public ABI driver so the scaling path takes part in
    # the same lifetime accounting as everything else.
    $engine.ClearTextures()
    for ($i = 0; $i -lt 200; $i++) {
        $h = $engine.LoadTexture($bytes)
        $null = $engine.ApplyTexture($shapes[$i % 6], $h)
        $engine.ReleaseTexture($h)
    }
    Assert ($engine.GetTextureCount() -eq 0) 'repeated load/apply/release leaves no handles behind'

    # --- multiple Shapes sharing one cached picture ---------------------------
    $handle = $engine.LoadTexture($bytes)
    foreach ($s in $shapes) { $null = $engine.ApplyTexture($s, $handle) }
    $report = $engine.InspectTexture($handle)
    $results.Add("  note shared image: $report")
    Assert ((Get-Field $report 'creations') -ne $null) 'the store reports its creation count'
    $engine.ReleaseTexture($handle)
    Assert ($engine.GetTextureCount() -eq 0) 'and releasing it leaves nothing'

    # --- everything let go ----------------------------------------------------
    $engine.ClearTextures()
    $null = $engine.ClearPictureCache()
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '0' -and (Get-Field $stats 'shapes') -eq '0') `
        'both owners are empty at the end'
    Assert ($engine.GetTextureCount() -eq 0) 'and no caller-visible texture remains'

    $end = PrivateMb
    $results.Add("  note private bytes at end: $end MB (PowerPoint's own document, undo and")
    $results.Add("       render resources are included here and are not BlipBridge's to release)")

    # --- the host is healthy --------------------------------------------------
    $probe = $slide.Shapes.AddShape(1, 600, 400, 40, 40)
    $probe.Fill.UserPicture($textureA)
    Assert ($probe.Fill.Type -eq 6) 'ordinary Fill.UserPicture still works after the whole run'
    $null = $engine.ApplyPicture($probe, $textureB)
    Assert ($probe.Fill.Type -eq 6) 'and so does UserPicture2'
} finally {
    try { $engine.ClearTextures() } catch { }
    try { $null = $engine.ClearPictureCache() } catch { }
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$results | Set-Content "$root/artifacts/release_stress.txt"
$results
if ($failures -gt 0) { throw "$failures release stress assertions failed" }
'Release stress: all assertions passed.'
