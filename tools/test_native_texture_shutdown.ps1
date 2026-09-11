<#
.SYNOPSIS
Leaves textures loaded, then quits PowerPoint.

.DESCRIPTION
The one lifetime boundary the main matrix cannot cover from inside a running
host: what happens when Office shuts down while texture handles are still held.
The Engine releases everything from OnDisconnection and from its destructor,
both of which run while GFX is still loaded, so a clean exit is the pass
condition. A hang, a crash dialog or a non-zero exit code is a failure.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

<#
 Runs one host to completion. $loadTextures decides whether textures are still
 held at Quit, so the same code path gives us a control: if a host with no
 textures also refuses to exit, the harness is at fault rather than the store.
#>
function Invoke-Host([bool]$loadTextures) {
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
    $hostId = $engine.GetHostProcessId()

    $presentation = $app.Presentations.Add(0)
    $slide = $presentation.Slides.Add(1, 12)
    $warmup = $slide.Shapes.AddShape(1, 400, 300, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    $loaded = 0
    if ($loadTextures) {
        [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
        $handles = @()
        foreach ($i in 1..5) { $handles += $engine.LoadTexture($bytes) }
        $shape = $slide.Shapes.AddShape(1, 20, 20, 100, 100)
        foreach ($h in $handles) { $null = $engine.ApplyTexture($shape, $h) }
        $loaded = $engine.GetTextureCount()
        $shape = $null
        $handles = $null
    }

    $presentation.Saved = -1
    $presentation.Close()

    # Every COM reference this scope holds must go before Quit, or PowerPoint
    # stays alive for reasons that have nothing to do with textures.
    $warmup = $null; $slide = $null; $presentation = $null
    $engine = $null; $addin = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $app.Quit()
    $app = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    return @{ Id = $hostId; Loaded = $loaded }
}

# Control first: no textures at all.
$control = Invoke-Host $false
$deadline = (Get-Date).AddSeconds(45)
while ((Get-Process -Id $control.Id -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 250
}
$controlSurvivor = Get-Process -Id $control.Id -ErrorAction SilentlyContinue
if ($controlSurvivor) {
    $controlSurvivor.Kill()
    throw 'Control host did not exit either; the harness, not the texture store, holds it open'
}
$results.Add('Control: a host that never loaded a texture exits cleanly.')

$session = Invoke-Host $true
$hostId = $session.Id
$results.Add("Loaded $($session.Loaded) textures and applied each; none released on purpose.")

$deadline = (Get-Date).AddSeconds(45)
while ((Get-Process -Id $hostId -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 250
}
$survivor = Get-Process -Id $hostId -ErrorAction SilentlyContinue
if ($survivor) {
    $survivor.Kill()
    throw 'PowerPoint did not exit with textures still loaded'
}
$results.Add('PowerPoint exited cleanly with five textures still loaded.')

# A fresh host must still work.
$again = Connect-PowerPoint
try {
    $again.COMAddIns.Update()
    $addin2 = $again.COMAddIns.Item('BlipBridge.Engine')
    $addin2.Connect = $true
    $engine2 = $addin2.Object
    if ($engine2.GetTextureCount() -ne 0) { throw 'A fresh host started with textures already loaded' }
    $p2 = $again.Presentations.Add(0)
    $s2 = $p2.Slides.Add(1, 12).Shapes.AddShape(1, 20, 20, 100, 100)
    $s2.Fill.UserPicture($texture)
    if ([int]$s2.Fill.Type -ne 6) { throw 'Fresh host cannot fill Shapes' }
    $p2.Saved = -1; $p2.Close()
    $results.Add('A fresh host starts with an empty store and works normally.')
} finally { $again.Quit() }

$results | Set-Content "$root/artifacts/native_texture_shutdown.txt"
$results
