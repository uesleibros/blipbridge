<#
.SYNOPSIS
Attributes any slowdown over a long run of applies.

.DESCRIPTION
A reported symptom for this class of workload is "fast at first, gradually
slower, eventually stabilises". Accepting "it stabilises" would leave the cause
unknown, so this measures apply time in buckets across a long run while tracking
everything that could plausibly grow underneath it:

  cached-image reference count   does document state accumulate per apply
  cached-image creations         is anything re-decoding behind our back
  texture handle count           is the store leaking
  private bytes                  is anything leaking natively

It runs the same loop twice: once letting the undo history accumulate, and once
calling StartNewUndoEntry so PowerPoint retires old entries. If the difference
between the two curves is the whole effect, the undo history is the cause.
#>
param([int]$Applies = 4000, [int]$Buckets = 8)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$hostProcess = Get-Process -Id $engine.GetHostProcessId()

function Get-PrivateMb { [math]::Round((Get-Process -Id $hostProcess.Id).PrivateMemorySize64 / 1MB, 1) }
function Get-Field([string]$report, [string]$name) {
    foreach ($part in $report.Split(';')) {
        $kv = $part.Split('=', 2)
        if ($kv.Length -eq 2 -and $kv[0] -eq $name) { return $kv[1] }
    }
    return $null
}

<#
 The third configuration is the one a renderer with changing content actually
 does: a fresh texture per frame rather than one reused texture. If growth comes
 from document image resources rather than from applying, this is where it shows.
#>
function Measure-Run([string]$label, [bool]$retireUndo, [bool]$freshTexturePerApply = $false) {
    $presentation = $app.Presentations.Add(0)
    try {
        $slide = $presentation.Slides.Add(1, 12)
        $warm = $slide.Shapes.AddShape(1, 500, 350, 40, 40)
        $warm.Fill.UserPicture($texture)
        $warm.Delete()

        [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
        $handle = $engine.LoadTexture($bytes)
        $shape = $slide.Shapes.AddShape(1, 20, 20, 120, 120)
        $perBucket = [int]($Applies / $Buckets)
        $baseMemory = Get-PrivateMb
        $results.Add("$label : $Applies applies in $Buckets buckets of $perBucket (retire undo: $retireUndo)")

        $first = 0.0
        for ($bucket = 0; $bucket -lt $Buckets; $bucket++) {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            for ($i = 0; $i -lt $perBucket; $i++) {
                if ($retireUndo) { $app.StartNewUndoEntry() }
                if ($freshTexturePerApply) {
                    $frame = $engine.LoadTexture($bytes)
                    $null = $engine.ApplyTexture($shape, $frame)
                    $engine.ReleaseTexture($frame)
                } else {
                    $null = $engine.ApplyTexture($shape, $handle)
                }
            }
            $watch.Stop()
            $per = $watch.Elapsed.TotalMilliseconds / $perBucket
            if ($bucket -eq 0) { $first = $per }
            $report = $engine.InspectTexture($handle)
            $drift = if ($first -gt 0) { [math]::Round($per / $first, 2) } else { 0 }
            $results.Add(("  bucket {0}: {1,7:N3} ms/apply  drift x{2,-5} refs={3,-5} creations={4,-4} private={5} MB" -f `
                $bucket, $per, $drift, (Get-Field $report 'cachedCount'), (Get-Field $report 'creations'), (Get-PrivateMb)))
        }
        $results.Add("  private bytes over the run: $baseMemory MB -> $(Get-PrivateMb) MB; handles=$($engine.GetTextureCount())")
        $engine.ReleaseTexture($handle)
    } finally {
        $presentation.Saved = -1
        $presentation.Close()
    }
}

Measure-Run 'one texture, undo accumulating' $false
Measure-Run 'one texture, undo retired    ' $true
Measure-Run 'new texture every apply      ' $false $true

$results | Set-Content "$root/artifacts/slowdown_attribution.txt"
$results
