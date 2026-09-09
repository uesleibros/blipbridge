<#
.SYNOPSIS
Checks whether a native apply writes any temporary image file.

.DESCRIPTION
MemoryImageToFill is the strictest of the three capabilities: it claims bytes in
memory reach a Shape fill with no temporary image file anywhere. Earlier tracing
showed Fill.UserPicture writes and re-reads a PNG under INetCache/Content.MSO,
which is exactly why the flag stayed false for the old MemoryFillExperiment.

Our own code opens no file, but that alone proves nothing about what Office does
internally. So this watches the Office image cache directories across two runs
with a distinctive image:

  1. a control run using Fill.UserPicture, which is expected to write there;
  2. a native ApplyTexture run, which should not.

The control matters: if the control writes nothing either, the watcher is
looking in the wrong place and the native result means nothing.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/BB_TRACE_UNIQUE_TEXTURE_01.png'
if (-not (Test-Path $texture)) { $texture = Join-Path $root 'artifacts/textures/texture_256_1.png' }
$results = New-Object System.Collections.Generic.List[string]

$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1
$msoFillPicture = 6

# Source copies live outside every watched root, so the control's own inputs are
# never mistaken for Office's cache output.
$sourceDirectory = Join-Path $root 'artifacts/no_temp_image_sources'
if (Test-Path $sourceDirectory) { Remove-Item $sourceDirectory -Recurse -Force }
New-Item -ItemType Directory -Path $sourceDirectory | Out-Null

# Office keeps decoded/temporary images under INetCache\Content.MSO; the exact
# root moved between Windows versions, so watch every plausible location.
$watchRoots = @(
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\INetCache\Content.MSO'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Temporary Internet Files\Content.MSO'),
    (Join-Path $env:TEMP '')
) | Where-Object { $_ -and (Test-Path $_) }

<#
 A before/after directory listing misses a file that Office creates and deletes
 between the two snapshots, which is the most likely reason the first version of
 this test saw nothing even for the control. FileSystemWatcher sees the create
 event itself, so short-lived cache files are caught.
#>
$script:seen = New-Object System.Collections.Generic.List[string]
$script:watchers = @()
foreach ($watch in $watchRoots) {
    $watcher = New-Object System.IO.FileSystemWatcher $watch
    $watcher.IncludeSubdirectories = $true
    $watcher.NotifyFilter = [IO.NotifyFilters]::FileName -bor [IO.NotifyFilters]::LastWrite
    $watcher.EnableRaisingEvents = $false
    $script:watchers += $watcher
    Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier "bbCreated$($script:watchers.Count)" | Out-Null
    Register-ObjectEvent -InputObject $watcher -EventName Changed -SourceIdentifier "bbChanged$($script:watchers.Count)" | Out-Null
}

function Start-Watching {
    Get-Event -SourceIdentifier bb* -ErrorAction SilentlyContinue | Remove-Event -ErrorAction SilentlyContinue
    foreach ($watcher in $script:watchers) { $watcher.EnableRaisingEvents = $true }
}
function Stop-Watching {
    Start-Sleep -Milliseconds 800
    foreach ($watcher in $script:watchers) { $watcher.EnableRaisingEvents = $false }
    $paths = @()
    foreach ($event in (Get-Event -SourceIdentifier bb* -ErrorAction SilentlyContinue)) {
        $paths += $event.SourceEventArgs.FullPath
        Remove-Event -EventIdentifier $event.EventIdentifier -ErrorAction SilentlyContinue
    }
    # Our own control copies live outside the watched roots, but filter anyway.
    return ($paths | Where-Object { $_ -notlike "$sourceDirectory*" } | Select-Object -Unique)
}

$results.Add("watching: $($watchRoots -join ' | ')")

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $warm = $slide.Shapes.AddShape($msoShapeRectangle, 500, 350, 40, 40)
    $warm.Fill.UserPicture($texture)
    $warm.Delete()
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)

    # ------------------------------------------------------------ control leg
    # Each control fill uses a *fresh source path*. Office caches by path, so
    # re-using one path after the warmup would hit the existing entry and write
    # nothing - which is exactly why the first version of this test was
    # inconclusive.
    Start-Watching
    for ($i = 0; $i -lt 12; $i++) {
        $copy = Join-Path $sourceDirectory ("bb_control_{0}.png" -f [guid]::NewGuid())
        Copy-Item -LiteralPath $texture -Destination $copy
        $shape = $slide.Shapes.AddShape($msoShapeRectangle, 20 + $i * 12, 20, 60, 60)
        $shape.Fill.UserPicture($copy)
        if ([int]$shape.Fill.Type -ne $msoFillPicture) { throw 'Control fill failed' }
    }
    $controlEvents = @(Stop-Watching)
    $controlNew = @($controlEvents | Where-Object { $_ -like '*Content.MSO*' })
    $results.Add("control (12x Fill.UserPicture): $($controlNew.Count) Content.MSO events, $($controlEvents.Count - $controlNew.Count) elsewhere")
    foreach ($file in ($controlNew | Select-Object -First 4)) { $results.Add("    $file") }

    # ------------------------------------------------------------- native leg
    $handle = $engine.LoadTexture($bytes)
    try {
        Start-Watching
        for ($i = 0; $i -lt 12; $i++) {
            $shape = $slide.Shapes.AddShape($msoShapeRectangle, 20 + $i * 12, 120, 60, 60)
            $null = $engine.ApplyTexture($shape, $handle)
            if ([int]$shape.Fill.Type -ne $msoFillPicture) { throw 'Native fill failed' }
        }
        $nativeEvents = @(Stop-Watching)
        $nativeNew = @($nativeEvents | Where-Object { $_ -like '*Content.MSO*' })
        $results.Add("native (12x ApplyTexture): $($nativeNew.Count) Content.MSO events, $($nativeEvents.Count - $nativeNew.Count) elsewhere (unrelated processes)")
        foreach ($file in ($nativeNew | Select-Object -First 4)) { $results.Add("    $file") }
    } finally {
        $engine.ReleaseTexture($handle)
    }

    $results.Add('')
    if ($controlNew.Count -eq 0) {
        $results.Add('INCONCLUSIVE: the control wrote nothing to Content.MSO either, so the')
        $results.Add('watcher is not seeing the cache. The native result proves nothing.')
    } elseif ($nativeNew.Count -eq 0) {
        $results.Add('PROVEN: Fill.UserPicture writes a temporary image per call under')
        $results.Add('Content.MSO; the native apply writes none. Bytes reach the fill with no')
        $results.Add('temporary image file.')
    } else {
        $results.Add('NOT PROVEN: the native apply also wrote under Content.MSO.')
    }
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

foreach ($watcher in $script:watchers) { $watcher.Dispose() }
Get-EventSubscriber -SourceIdentifier bb* -ErrorAction SilentlyContinue | Unregister-Event -ErrorAction SilentlyContinue

$results | Set-Content "$root/artifacts/no_temp_image.txt"
$results
