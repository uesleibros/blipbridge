<#
.SYNOPSIS
Decodes image bytes into an Office cached image with no file on disk.

.DESCRIPTION
Exercises LoadCachedImageExperiment, the smallest native step toward a reusable
texture. It calls only the GFX export ICachedImage::Create, resolved by name,
and releases everything it creates. No Shape, slide or presentation is touched,
so a failure cannot damage a document.

An ordinary picture fill runs first purely to make sure GFX is loaded; the
experiment refuses to run otherwise rather than loading a second copy itself.
The same bytes are then decoded several times, and different images once each,
so the report shows whether Office hands back a distinct cached image per call.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$results = New-Object System.Collections.Generic.List[string]

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    # Warm GFX. This is the only picture fill in the test and it is not measured.
    $warmup = $presentation.Slides.Add(1, 12).Shapes.AddShape(1, 10, 10, 60, 60)
    $warmup.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_32_0.png'))
    $warmup.Delete()

    function Get-Field([string]$report, [string]$name) {
        foreach ($part in $report.Split(';')) {
            $pair = $part.Split('=', 2)
            if ($pair.Length -eq 2 -and $pair[0] -eq $name) { return $pair[1] }
        }
        return $null
    }

    $addresses = New-Object System.Collections.Generic.List[string]
    foreach ($file in 'texture_64_1.png', 'texture_64_1.png', 'texture_64_1.png',
                      'texture_256_1.png', 'texture_64_0.jpg') {
        [byte[]]$bytes = [IO.File]::ReadAllBytes("$root/artifacts/textures/$file")
        $report = $engine.LoadCachedImageExperiment($bytes)
        $results.Add("$file ($($bytes.Length) bytes): $report")
        if ((Get-Field $report 'cachedCount') -ne '1') {
            throw "Cached image arrived with an unexpected reference count for $file"
        }
        $addresses.Add((Get-Field $report 'cached'))
    }

    # Every cached image was released before the next call, so address reuse is
    # expected and proves nothing either way. Only the successful decode matters.
    $results.Add("Distinct cached addresses observed: $(($addresses | Select-Object -Unique).Count) of $($addresses.Count)")
    $results.Add('Decoded PNG and JPEG bytes from memory with no source file.')

    try {
        [byte[]]$garbage = 1..64 | ForEach-Object { [byte]($_ % 256) }
        $report = $engine.LoadCachedImageExperiment($garbage)
        $results.Add("Invalid bytes unexpectedly accepted: $report")
    } catch {
        $results.Add('Invalid bytes rejected without a crash.')
    }

    # The host must still be healthy afterwards.
    $check = $presentation.Slides.Add(2, 12).Shapes.AddShape(1, 10, 10, 60, 60)
    $check.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_32_0.png'))
    if ([int]$check.Fill.Type -ne 6) { throw 'PowerPoint fill broke after the experiment' }
    $results.Add('Ordinary UserPicture still works after the native calls.')
} finally {
    $results | Set-Content "$root/artifacts/cached_image_load.txt"
    $results
    $presentation.Saved = -1
    $presentation.Close()
}
