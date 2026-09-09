$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

function Assert-Rejected {
    param([scriptblock]$Operation, [string]$Description)

    $rejected = $false
    try {
        & $Operation | Out-Null
    } catch {
        $rejected = $true
    }
    if (-not $rejected) {
        throw "Expected rejection: $Description"
    }
}

# Uses only the normal donor fallback. No memory adapter or private Office calls.
$application = New-Object -ComObject PowerPoint.Application
$application.COMAddIns.Update()
$addin = $application.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $application.Presentations.Add(0)

try {
    $slide = $presentation.Slides.Add(1, 12)
    $donor = $slide.Shapes.AddShape(1, 10, 10, 100, 100)
    $destination = $slide.Shapes.AddShape(1, 130, 10, 100, 100)
    $textBox = $slide.Shapes.AddTextbox(1, 10, 140, 100, 100)

    Assert-Rejected { $engine.RegisterTextureShape($donor) } 'donor without picture fill'
    Assert-Rejected { $engine.RegisterTextureShape($textBox) } 'unsupported donor type'

    $donor.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_64_0.png'))
    $handle = $engine.RegisterTextureShape($donor)
    Assert-Rejected { $engine.ApplyTexture($textBox, $handle) } 'unsupported destination type'
    Assert-Rejected { $engine.ApplyTexture($destination, -1) } 'unknown handle'
    Assert-Rejected { $engine.ApplyTexture($destination) } 'missing handle argument'

    $engine.ReleaseTexture($handle)
    Assert-Rejected { $engine.ReleaseTexture($handle) } 'double release'
    $engine.ClearTextures()
    $engine.ClearTextures()
    if ($engine.GetTextureCount() -ne 0) {
        throw 'Repeated ClearTextures must leave an empty cache'
    }

    # Capabilities are now a statement about this host rather than a constant.
    # On a validated Office build inside PowerPoint the native backend is
    # available and the flags read True; anywhere else they must read False.
    # Either answer is correct - what must hold is that they agree with whether
    # LoadTexture actually works.
    $capabilities = $engine.GetCapabilities()
    $claimsNative = $capabilities -match 'CachedTextureApply=True'
    [byte[]]$probeBytes = [IO.File]::ReadAllBytes((Join-Path $root 'artifacts/textures/texture_32_0.png'))
    $nativeWorks = $false
    try {
        $probeHandle = $engine.LoadTexture($probeBytes)
        $engine.ReleaseTexture($probeHandle)
        $nativeWorks = $true
    } catch { $nativeWorks = $false }
    if ($claimsNative -ne $nativeWorks) {
        throw "Capabilities claim CachedTextureApply=$claimsNative but LoadTexture works=$nativeWorks"
    }
    if ($capabilities -notmatch 'PickUpFallback=True') {
        throw 'The donor fallback must stay advertised'
    }
    if ($claimsNative -and $capabilities -notmatch 'MemoryImageToFill=True') {
        throw 'Native backend available but MemoryImageToFill not advertised'
    }
    'Fallback contract tests passed: invalid types/handles, argument count, double release, clear, capabilities.' |
        Tee-Object (Join-Path $root 'artifacts/fallback_contract.txt')
} finally {
    $engine.ClearTextures()
    $presentation.Saved = -1
    $presentation.Close()
}
