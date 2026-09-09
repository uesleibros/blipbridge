<#
.SYNOPSIS
Compares encoded-image loading against raw-pixel loading across sizes.

.DESCRIPTION
Both legs run in process through the public C ABI. The encoded leg reads a real
PNG; the raw leg hands over a BGRA buffer of the same dimensions. Each texture is
released immediately, so what is timed is creation and teardown.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$results = New-Object System.Collections.Generic.List[string]

<#
 Only sizes where a PNG of the *same* dimensions exists are comparable. Pairing a
 256x256 PNG against a 640x480 raw buffer would flatter the encoded leg by giving
 it a sixth of the pixels, so those rows are marked and read as raw-path scaling
 rather than as a comparison.
#>
$cases = @(
    @{ Png = 'texture_32_0.png';  W = 32;  H = 32;  Matched = $true },
    @{ Png = 'texture_64_1.png';  W = 64;  H = 64;  Matched = $true },
    @{ Png = 'texture_128_0.png'; W = 128; H = 128; Matched = $true },
    @{ Png = 'texture_256_1.png'; W = 256; H = 256; Matched = $true },
    @{ Png = 'texture_256_1.png'; W = 320; H = 288; Matched = $false },
    @{ Png = 'texture_256_1.png'; W = 640; H = 480; Matched = $false }
)

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, 12)
    $warm = $slide.Shapes.AddShape(1, 500, 350, 40, 40)
    $warm.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_64_1.png'))
    $warm.Delete()

    foreach ($case in $cases) {
        $png = Join-Path $root ('artifacts/textures/' + $case.Png)
        $tag = if ($case.Matched) { 'comparable=1' } else { 'comparable=0(encoded leg is a different size)' }
        $results.Add("$tag;" + $engine.BenchmarkPixelLoad($png, $case.W, $case.H, 200))
    }
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

$results | Set-Content "$root/artifacts/pixel_benchmark.txt"
$results
