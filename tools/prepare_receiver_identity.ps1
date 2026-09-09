<#
.SYNOPSIS
Prepares three ordinary AutoShapes and fills each with the same picture.

.DESCRIPTION
Companion to experiments/exp_internal_blip/probe_receiver_identity.py. Two
shapes live on the first slide and one on a second slide, and the first shape is
filled twice, so the probe can tell whether the OART receiver used by the fill
transaction is per call, per Shape, per slide or per document.

This script deliberately calls the public Fill.UserPicture directly rather than
the instrumented research entry point: no IAT hooks are installed, so the
observed dispatch is the ordinary Office path.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
try {
    $first = $presentation.Slides.Add(1, 12)
    $second = $presentation.Slides.Add(2, 12)
    $shapes = @(
        $first.Shapes.AddShape(1, 10, 10, 100, 100),
        $first.Shapes.AddShape(1, 140, 10, 100, 100),
        $second.Shapes.AddShape(1, 10, 10, 100, 100)
    )
    for ($index = 0; $index -lt $shapes.Count; $index++) {
        $shapes[$index].Name = "BB_receiver_$index"
    }

    # Warm the decoder and OART modules before the module snapshot is taken.
    $warmup = $first.Shapes.AddShape(1, 270, 10, 60, 60)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    $identityBefore = $shapes | ForEach-Object {
        @($_.Id, $_.Name, $_.Type, $_.Left, $_.Top, $_.Width, $_.Height,
          $_.Rotation, $_.ZOrderPosition) -join '|'
    }

    $process = Get-Process -Id $engine.GetHostProcessId()
    $moduleList = @($process.Modules | ForEach-Object {
        @{
            name = $_.ModuleName
            base = $_.BaseAddress.ToInt64()
            size = $_.ModuleMemorySize
            version = $_.FileVersionInfo.FileVersion
        }
    })
    @{ pid = $process.Id; modules = $moduleList } |
        ConvertTo-Json -Depth 4 |
        Set-Content "$root/artifacts/decoder_target.json"

    'Ready for debugger trigger.'
    $trigger = "$root/artifacts/decoder_go"
    $deadline = (Get-Date).AddSeconds(55)
    while (!(Test-Path -LiteralPath $trigger)) {
        if ((Get-Date) -gt $deadline) { throw 'Debugger trigger timed out' }
        Start-Sleep -Milliseconds 200
    }

    # Call order is A, A, B, C. The repeat on A is what separates "per Shape"
    # from "per call": anything that changes between the first two calls cannot
    # be Shape identity. Same bytes throughout, so nothing varies from the image.
    foreach ($shape in @($shapes[0], $shapes[0], $shapes[1], $shapes[2])) {
        $shape.Fill.UserPicture($texture)
    }

    $identityAfter = $shapes | ForEach-Object {
        @($_.Id, $_.Name, $_.Type, $_.Left, $_.Top, $_.Width, $_.Height,
          $_.Rotation, $_.ZOrderPosition) -join '|'
    }
    for ($index = 0; $index -lt $shapes.Count; $index++) {
        if ($identityAfter[$index] -ne $identityBefore[$index]) {
            throw "Shape $index identity or geometry changed"
        }
        if ([int]$shapes[$index].Fill.Type -ne 6) {
            throw "Shape $index does not have a picture fill"
        }
    }
    ($shapes | ForEach-Object { "Shape $($_.Name) id=$($_.Id) preserved with picture fill" }) |
        Set-Content "$root/artifacts/receiver_identity_validation.txt"
    'Receiver identity trace completed.'
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}
