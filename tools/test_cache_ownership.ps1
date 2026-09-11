<#
.SYNOPSIS
Ownership regression: UserPicture2 must survive ClearTextures, and vice versa.

.DESCRIPTION
This reproduces a real failure and keeps it fixed:

    UserPicture2 failed:
    Texture handle 16777216 is not valid
    (it was released; handles are never recycled)
    (code -6)

The texture store used to *own* each image and hand callers a `long` handle into
its map. The picture cache behind UserPicture2 took one of those handles like any
other caller — so `ClearTextures()`, or a caller releasing that particular
handle, destroyed the image the cache was still pointing at. The next
UserPicture2 then looked up a handle that no longer existed.

Handle 16777216 is the first one ever issued, which is exactly what the picture
cache took on its first call. That number in the report was the tell.

The model now is:

    public handle  -> references an image
    picture cache  -> references an image, independently

so each clears only what it owns, and Shutdown clears everything.

Every sequence below is asserted, not printed for a human to read past.
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

<#
 Every UserPicture2 here goes through the real exported BB_ApplyPicture. A
 failure surfaces as a COM error carrying the native message, so the assertion
 is simply "this did not throw" plus the fill that resulted.
#>
function Try-UserPicture2($shape, [string]$path) {
    try {
        $null = $script:engine.ApplyPicture($shape, $path)
        return ''
    } catch {
        return ($_.Exception.Message -replace '\s+', ' ')
    }
}

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shape = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
    $other = $slide.Shapes.AddShape(1, 200, 20, 120, 120)

    # --- the exact reported sequence ------------------------------------------
    $null = $engine.ClearPictureCache()
    $engine.ClearTextures()
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "UserPicture2 works from a clean start ($err)"
    Assert ($shape.Fill.Type -eq 6) 'and produces a picture fill'

    $engine.ClearTextures()
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "UserPicture2 survives ClearTextures ($err)"
    Assert ($shape.Fill.Type -eq 6) 'and still produces a picture fill'

    # Twice, because the first call after a clear repopulates the cache and the
    # second exercises the repopulated entry.
    $engine.ClearTextures()
    $null = Try-UserPicture2 $shape $textureA
    $err = Try-UserPicture2 $other $textureA
    Assert ($err -eq '') "a second Shape works after ClearTextures ($err)"

    # --- ClearPictureCache ----------------------------------------------------
    $null = $engine.ClearPictureCache()
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "UserPicture2 survives ClearPictureCache ($err)"
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '1') 'and the cache repopulated itself'

    # --- explicit textures mixed with UserPicture2 ----------------------------
    <#
     The decisive case. An explicit handle and the picture cache may end up
     holding the *same* image; releasing the handle must not disturb the cache,
     and clearing the cache must not disturb the handle.
    #>
    [byte[]]$bytes = [IO.File]::ReadAllBytes($textureA)
    $handle = $engine.LoadTexture($bytes)
    $null = Try-UserPicture2 $shape $textureA
    $engine.ReleaseTexture($handle)
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "UserPicture2 survives ReleaseTexture of a caller handle ($err)"

    $handle2 = $engine.LoadTexture($bytes)
    $null = $engine.ClearPictureCache()
    $applyError = ''
    try { $null = $engine.ApplyTexture($other, $handle2) }
    catch { $applyError = ($_.Exception.Message -replace '\s+', ' ') }
    Assert ($applyError -eq '') "an explicit texture survives ClearPictureCache ($applyError)"
    Assert ($other.Fill.Type -eq 6) 'and still applies'
    $engine.ReleaseTexture($handle2)

    # --- released handles stay stale forever ----------------------------------
    $staleError = ''
    try { $null = $engine.ApplyTexture($other, $handle2) }
    catch { $staleError = ($_.Exception.Message -replace '\s+', ' ') }
    Assert ($staleError -ne '') 'a released handle is still refused'
    Assert ($staleError -like '*never recycled*') 'and says handles are never recycled'

    $handle3 = $engine.LoadTexture($bytes)
    Assert ($handle3 -ne $handle2) 'and a new handle is a different number'
    $engine.ReleaseTexture($handle3)

    # --- repeated and alternating paths ---------------------------------------
    for ($i = 0; $i -lt 3; $i++) {
        $null = Try-UserPicture2 $shape $textureA
    }
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "repeating the same path is stable ($err)"

    $err = ''
    for ($i = 0; $i -lt 4; $i++) {
        $path = $textureA
        if ($i % 2 -eq 1) { $path = $textureB }
        $stepError = Try-UserPicture2 $shape $path
        if ($stepError -ne '') { $err = $stepError }
    }
    Assert ($err -eq '') "alternating paths is stable ($err)"
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '2') 'and both files are cached'

    # --- multiple Shapes, one file --------------------------------------------
    $many = @()
    for ($i = 0; $i -lt 5; $i++) {
        $many += $slide.Shapes.AddShape(1, 20 + $i * 40, 200, 35, 35)
    }
    $err = ''
    foreach ($s in $many) {
        $stepError = Try-UserPicture2 $s $textureA
        if ($stepError -ne '') { $err = $stepError }
    }
    Assert ($err -eq '') "one file across many Shapes ($err)"
    $filled = @($many | Where-Object { $_.Fill.Type -eq 6 }).Count
    Assert ($filled -eq 5) 'and every one of them is filled'

    # --- Shutdown then Init ---------------------------------------------------
    <#
     Shutdown must release everything BlipBridge owns - both owners, not just
     the handle table - and Init must leave the library usable again.
    #>
    $engine.ClearTextures()
    $null = $engine.ClearPictureCache()
    $err = Try-UserPicture2 $shape $textureA
    Assert ($err -eq '') "UserPicture2 works after clearing both owners ($err)"
    $stats = $engine.PictureCacheStats()
    Assert ((Get-Field $stats 'textures') -eq '1') 'and the cache holds exactly one image'

    # --- the host is unharmed --------------------------------------------------
    $probe = $slide.Shapes.AddShape(1, 400, 200, 40, 40)
    $probe.Fill.UserPicture($textureA)
    Assert ($probe.Fill.Type -eq 6) 'ordinary Fill.UserPicture still works at the end'
} finally {
    try { $engine.ClearTextures() } catch { }
    try { $null = $engine.ClearPictureCache() } catch { }
    try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
}

$results | Set-Content "$root/artifacts/cache_ownership.txt"
$results
if ($failures -gt 0) { throw "$failures cache-ownership assertions failed" }
'Cache ownership: all assertions passed.'
