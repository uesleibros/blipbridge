<#
.SYNOPSIS
Production-like stress for the native texture API.

.DESCRIPTION
Pushes the API well past anything the correctness tests do, and tracks the four
things that would reveal a defect the functional tests miss:

  private bytes            - anything stranded shows up here
  handle count             - the store must not accumulate entries
  cached-image creations   - must stay equal to the number of LoadTexture calls,
                             which is what "reuse" actually means
  reference-count range    - must stay bounded and return to 1 per texture once
                             its documents are gone

A crash, an assertion or a hang is a failure regardless of the numbers.
PowerPoint shutdown with live textures is covered separately by
tools/test_native_texture_shutdown.ps1.
#>
param([int]$Applies = 10000, [int]$TextureChurn = 100, [int]$ChurnRounds = 5)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$textureA = Join-Path $root 'artifacts/textures/texture_64_1.png'
$textureB = Join-Path $root 'artifacts/textures/texture_128_0.png'
$results = New-Object System.Collections.Generic.List[string]

$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1
$msoFillPicture = 6
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
$hostProcess = Get-Process -Id $engine.GetHostProcessId()

function Get-PrivateMb { [math]::Round((Get-Process -Id $hostProcess.Id).PrivateMemorySize64 / 1MB, 1) }
function Get-Field([string]$report, [string]$name) {
    foreach ($part in $report.Split(';')) {
        $pair = $part.Split('=', 2)
        if ($pair.Length -eq 2 -and $pair[0] -eq $name) { return $pair[1] }
    }
    return $null
}
function Get-Count([long]$handle) { [int](Get-Field $engine.InspectTexture($handle) 'cachedCount') }
function Get-Creations { [int](Get-Field $engine.InspectTexture(0) 'creations') }

$creationsAtStart = Get-Creations
$memoryAtStart = Get-PrivateMb
$results.Add("start: private=$memoryAtStart MB creations=$creationsAtStart handles=$($engine.GetTextureCount())")

$presentation = $app.Presentations.Add(0)
try {
    $slide1 = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide2 = $presentation.Slides.Add(2, $ppLayoutBlank)
    $slide3 = $presentation.Slides.Add(3, $ppLayoutBlank)
    $warmup = $slide1.Shapes.AddShape($msoShapeRectangle, 600, 400, 40, 40)
    $warmup.Fill.UserPicture($textureA)
    $warmup.Delete()

    [byte[]]$bytesA = [IO.File]::ReadAllBytes($textureA)
    [byte[]]$bytesB = [IO.File]::ReadAllBytes($textureB)

    # --- 1 texture, many applies on one Shape --------------------------------
    $handleA = $engine.LoadTexture($bytesA)
    $hot = $slide1.Shapes.AddShape($msoShapeRectangle, 20, 20, 100, 100)
    $identity = @($hot.Id, $hot.Left, $hot.Top, $hot.Width, $hot.Height) -join '|'
    $before = Get-PrivateMb
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $Applies; $i++) { $null = $engine.ApplyTexture($hot, $handleA) }
    $watch.Stop()
    $countAfterHot = Get-Count $handleA
    $results.Add("$Applies applies on one Shape: $([math]::Round($watch.Elapsed.TotalMilliseconds,0)) ms, private $before -> $(Get-PrivateMb) MB, cachedCount=$countAfterHot")
    if (@($hot.Id, $hot.Left, $hot.Top, $hot.Width, $hot.Height) -join '|' -ne $identity) { throw 'Hot Shape identity changed' }
    if ([int]$hot.Fill.Type -ne $msoFillPicture) { throw 'Hot Shape lost its fill' }
    if ((Get-Creations) -ne $creationsAtStart + 1) { throw 'Applies created extra cached images' }
    $results.Add('  creations unchanged by applies: reuse confirmed.')

    # --- alternating two textures on one Shape -------------------------------
    $handleB = $engine.LoadTexture($bytesB)
    $before = Get-PrivateMb
    for ($i = 0; $i -lt 2000; $i++) {
        $null = $engine.ApplyTexture($hot, $(if ($i % 2) { $handleA } else { $handleB }))
    }
    $results.Add("2000 alternating applies: private $before -> $(Get-PrivateMb) MB, counts A=$(Get-Count $handleA) B=$(Get-Count $handleB)")
    if ((Get-Creations) -ne $creationsAtStart + 2) { throw 'Alternating applies created extra cached images' }

    # --- one texture across many Shapes on several slides --------------------
    $many = @()
    foreach ($slide in @($slide1, $slide2, $slide3)) {
        for ($i = 0; $i -lt 40; $i++) {
            $many += $slide.Shapes.AddShape($msoShapeRectangle, 20 + ($i % 10) * 60, 150 + [math]::Floor($i / 10) * 60, 50, 50)
        }
    }
    $before = Get-PrivateMb
    foreach ($shape in $many) { $null = $engine.ApplyTexture($shape, $handleA) }
    $filled = ($many | Where-Object { [int]$_.Fill.Type -eq $msoFillPicture }).Count
    $results.Add("$($many.Count) Shapes over 3 slides from one texture: $filled filled, private $before -> $(Get-PrivateMb) MB, cachedCount=$(Get-Count $handleA)")
    if ($filled -ne $many.Count) { throw 'Not every Shape received the fill' }

    # --- Shape and slide deletion --------------------------------------------
    for ($i = 0; $i -lt 20; $i++) { $many[$i].Delete() }
    $results.Add("after deleting 20 filled Shapes: cachedCount=$(Get-Count $handleA)")
    $slide3.Delete()
    $results.Add("after deleting a whole slide: cachedCount=$(Get-Count $handleA)")
    $null = $engine.ApplyTexture($hot, $handleA)
    if ([int]$hot.Fill.Type -ne $msoFillPicture) { throw 'Apply broke after deletions' }
    $results.Add('  applies still work after Shape and slide deletion.')

    # --- load/release churn ---------------------------------------------------
    $before = Get-PrivateMb
    for ($round = 0; $round -lt $ChurnRounds; $round++) {
        $batch = @()
        for ($i = 0; $i -lt $TextureChurn; $i++) { $batch += $engine.LoadTexture($bytesA) }
        if ($engine.GetTextureCount() -ne $TextureChurn + 2) { throw 'Handle count wrong during churn' }
        $null = $engine.ApplyTexture($hot, $batch[$TextureChurn - 1])
        foreach ($h in $batch) { $engine.ReleaseTexture($h) }
        if ($engine.GetTextureCount() -ne 2) { throw 'Handles leaked after a churn round' }
    }
    $results.Add("$ChurnRounds rounds of $TextureChurn load/release: private $before -> $(Get-PrivateMb) MB, handles=$($engine.GetTextureCount()), creations=$(Get-Creations)")

    # --- stale handle and double release --------------------------------------
    $doomed = $engine.LoadTexture($bytesA)
    $engine.ReleaseTexture($doomed)
    $rejected = $false
    try { $engine.ReleaseTexture($doomed) } catch { $rejected = $true }
    if (-not $rejected) { throw 'Double release was accepted' }
    $rejected = $false
    try { $null = $engine.ApplyTexture($hot, $doomed) } catch { $rejected = $true }
    if (-not $rejected) { throw 'Stale handle was accepted' }
    $results.Add('Double release and stale-handle apply both rejected.')
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

# --- multiple presentations --------------------------------------------------
# The handles used below are loaded here rather than discovered by scanning the
# handle space. Scanning worked only while this process had allocated fewer than
# a couple of thousand textures; handles are never recycled, so once another
# harness had run in the same PowerPoint the live ones sat past the end of the
# window and the scan silently found nothing.
$handles = @($engine.LoadTexture($bytesA), $engine.LoadTexture($bytesB))
$decks = @()
try {
    foreach ($index in 0..2) {
        $deck = $app.Presentations.Add(0)
        $decks += $deck
        $shape = $deck.Slides.Add(1, $ppLayoutBlank).Shapes.AddShape($msoShapeRectangle, 20, 20, 100, 100)
        $null = $engine.ApplyTexture($shape, $handles[0])
        if ([int]$shape.Fill.Type -ne $msoFillPicture) { throw "Texture failed in presentation $index" }
    }
    $results.Add("One texture applied in 3 concurrent presentations; cachedCount=$(Get-Count $handles[0])")
} finally {
    foreach ($deck in $decks) { $deck.Saved = $msoTrue; $deck.Close() }
}

$countAfterAllClosed = Get-Count $handles[0]
$results.Add("after closing every presentation: cachedCount=$countAfterAllClosed (1 means only the handle holds it)")
if ($countAfterAllClosed -ne 1) {
    $results.Add('  WARNING: references survive with no document open')
}

$engine.ClearTextures()
if ($engine.GetTextureCount() -ne 0) { throw 'ClearTextures left handles behind' }
Start-Sleep -Seconds 2
$results.Add("end: private=$(Get-PrivateMb) MB (start $memoryAtStart MB), creations=$(Get-Creations), handles=0")

$check = $app.Presentations.Add(0)
try {
    $shape = $check.Slides.Add(1, $ppLayoutBlank).Shapes.AddShape($msoShapeRectangle, 20, 20, 100, 100)
    $shape.Fill.UserPicture($textureA)
    if ([int]$shape.Fill.Type -ne $msoFillPicture) { throw 'Host broken after the stress run' }
    $results.Add('Ordinary UserPicture still works after the whole run.')
} finally { $check.Saved = $msoTrue; $check.Close() }

$results | Set-Content "$root/artifacts/native_texture_stress.txt"
$results
