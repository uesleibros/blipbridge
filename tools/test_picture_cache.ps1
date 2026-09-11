<#
.SYNOPSIS
Verifies BB_ApplyPicture: the dispatch decision, both caches, and the errors.

.DESCRIPTION
Runs the real C ABI in process inside PowerPoint, through a small driver the
Engine exposes, and asserts the behaviour the documentation promises:

  dispatch   a native class takes the native path; a class without one falls
             back to Fill.UserPicture; a class with neither is refused by name
  path cache the same file applied to many Shapes is decoded once
  skip cache the same image on the same Shape twice does no Office work
  freshness  editing the file on disk produces the new image, not the old
  safety     a fill changed behind the cache's back is re-applied, not skipped
  errors     each failure reports its own reason, never a generic one

Everything is asserted, not printed for a human to eyeball: a wrong answer fails
the script.
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
$ppLayoutBlank = 12

$textureA = Join-Path $root 'artifacts/textures/texture_128_0.png'
$textureB = Join-Path $root 'artifacts/textures/texture_64_1.png'
if (-not (Test-Path $textureA)) { throw "Missing $textureA - run tools/generate_textures.ps1" }

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


$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    # --- dispatch: a native class ---------------------------------------------
    $auto = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
    $null = $engine.ClearPictureCache()
    $null = $engine.ApplyPicture($auto, $textureA)
    Assert ($auto.Fill.Type -eq 6) 'AutoShape gets a picture fill'
    Assert ($auto.Type -eq 1) 'AutoShape keeps its type'
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '1') 'one texture cached after one file'

    # --- path cache: many Shapes, one decode ----------------------------------
    $others = @()
    for ($i = 0; $i -lt 5; $i++) {
        $others += $slide.Shapes.AddShape(1, 200 + $i * 30, 20, 25, 25)
    }
    foreach ($shape in $others) { $null = $engine.ApplyPicture($shape, $textureA) }
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '1') 'five more Shapes decode the file no further times'
    Assert ((Get-Field $stats 'shapes') -eq '6') 'each Shape is remembered'

    # --- skip cache: the same image on the same Shape --------------------------
    $before = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    $null = $engine.ApplyPicture($auto, $textureA)
    $null = $engine.ApplyPicture($auto, $textureA)
    $after = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    Assert (($after - $before) -eq 2) 'repeating the same image on the same Shape is skipped'

    # A different image on the same Shape must not be skipped.
    $before = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    $null = $engine.ApplyPicture($auto, $textureB)
    $after = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    Assert (($after - $before) -eq 0) 'a different image on the same Shape does real work'
    Assert ((Get-Field ($engine.PictureCacheStats()) 'textures') -eq '2') 'the second file is cached too'

    # --- safety: a fill changed behind the cache's back ------------------------
    $null = $engine.ApplyPicture($auto, $textureA)
    $auto.Fill.Solid()                      # no longer a picture fill at all
    $before = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    $null = $engine.ApplyPicture($auto, $textureA)
    $after = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    Assert (($after - $before) -eq 0) 'a fill replaced by a solid colour is re-applied, not skipped'
    Assert ($auto.Fill.Type -eq 6) 'and the picture fill is back'

    # --- freshness: the file changes on disk ----------------------------------
    $scratch = Join-Path $env:TEMP ('bbcache_' + [guid]::NewGuid().ToString('N') + '.png')
    Copy-Item $textureA $scratch
    $fresh = $slide.Shapes.AddShape(1, 20, 200, 100, 100)
    $null = $engine.ApplyPicture($fresh, $scratch)
    $texturesBefore = [int](Get-Field ($engine.PictureCacheStats()) 'textures')
    Start-Sleep -Milliseconds 1100          # so the last-write time is visibly different
    Copy-Item $textureB $scratch -Force
    $before = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    $null = $engine.ApplyPicture($fresh, $scratch)
    $after = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    Assert (($after - $before) -eq 0) 'a rewritten file is not served from the cache'
    Assert ([int](Get-Field ($engine.PictureCacheStats()) 'textures') -eq $texturesBefore) `
        'and the superseded texture is released rather than accumulating'
    Remove-Item $scratch -ErrorAction SilentlyContinue

    # --- invalidation ---------------------------------------------------------
    $null = $engine.ApplyPicture($auto, $textureA)
    $null = $engine.InvalidateShape($auto)
    $before = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    $null = $engine.ApplyPicture($auto, $textureA)
    $after = [int](Get-Field ($engine.PictureCacheStats()) 'skipped')
    Assert (($after - $before) -eq 0) 'InvalidateShape makes the next apply do real work'

    # --- dispatch: a class with no native path --------------------------------
    $table = $slide.Shapes.AddTable(2, 2, 300, 200, 180, 70)
    $null = $engine.ApplyPicture($table, $textureA)
    Assert ($table.Fill.Type -eq 6) 'a Table falls back to Fill.UserPicture and is filled'
    Assert ($table.Type -eq 19) 'and keeps its type'

    # --- dispatch: a class with neither ---------------------------------------
    $connector = $slide.Shapes.AddConnector(1, 500, 300, 600, 380)
    $connectorError = ''
    try { $null = $engine.ApplyPicture($connector, $textureA) }
    catch { $connectorError = $_.Exception.Message }
    Assert ($connectorError -ne '') 'a Connector is refused rather than filled'
    Assert ($connectorError -notlike '*apply failed*') 'and the refusal is specific, not generic'
    $alive = $false
    try { $null = $app.Version; $alive = $true } catch { }
    Assert $alive 'and PowerPoint is still running'

    # --- errors: a path that is not there -------------------------------------
    $missingError = ''
    try { $null = $engine.ApplyPicture($auto, (Join-Path $env:TEMP 'no_such_image_12345.png')) }
    catch { $missingError = $_.Exception.Message }
    Assert ($missingError -like '*image file*') 'a missing file reports itself as a missing file'

    # --- clearing --------------------------------------------------------------
    $null = $engine.ClearPictureCache()
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '0') 'ClearPictureCache releases every texture'
    Assert ((Get-Field $stats 'shapes') -eq '0') 'and forgets every Shape'

    <#
     What the skip is actually worth. Both legs apply the same image to the same
     Shape; the first has the cache primed, the second invalidates before each
     call so every one does real work.
    #>
    $measured = $slide.Shapes.AddShape(1, 600, 20, 80, 80)
    $null = $engine.ApplyPicture($measured, $textureA)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt 200; $i++) { $null = $engine.ApplyPicture($measured, $textureA) }
    $watch.Stop()
    $skippedMs = $watch.Elapsed.TotalMilliseconds / 200
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt 200; $i++) {
        $null = $engine.InvalidateShape($measured)
        $null = $engine.ApplyPicture($measured, $textureA)
    }
    $watch.Stop()
    # The invalidate call is itself a COM round trip, so it is measured and
    # subtracted rather than being quietly charged to the apply.
    $watch2 = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt 200; $i++) { $null = $engine.InvalidateShape($measured) }
    $watch2.Stop()
    $workingMs = ($watch.Elapsed.TotalMilliseconds - $watch2.Elapsed.TotalMilliseconds) / 200
    $results.Add(("  note skipped apply {0:N4} ms vs real apply {1:N4} ms (both include one COM round trip)" -f `
        $skippedMs, $workingMs))
    Assert ($skippedMs -lt $workingMs) 'a skipped apply is cheaper than a real one'

    # --- the host is unharmed --------------------------------------------------
    $after = $slide.Shapes.AddShape(1, 400, 20, 60, 60)
    $after.Fill.UserPicture($textureA)
    Assert ($after.Fill.Type -eq 6) 'ordinary Fill.UserPicture still works afterwards'
} finally {
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$results | Set-Content "$root/artifacts/picture_cache.txt"
$results
if ($failures -gt 0) { throw "$failures picture-cache assertions failed" }
'Picture cache: all assertions passed.'
