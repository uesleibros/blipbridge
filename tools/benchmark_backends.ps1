<#
.SYNOPSIS
Measures what each backend actually costs, against Office's own API as the
control.

.DESCRIPTION
The portable backend is not fast and this exists to say how much. Its value is
compatibility - the same API, through documented Office calls, where the
accelerated path is unavailable - and a number is the only honest way to describe
the trade.

Every measurement here is one real Office edit on a real Shape in a real
presentation. The control is `Shape.Fill.UserPicture` called directly from
PowerShell, which is what a caller would write without BlipBridge at all.

Run it once against a normal x64 build and once against one configured with
BB_FORCE_PORTABLE_BACKEND=ON. It records which backend answered, so the two runs
can be put side by side.

## Reading the results

**Every per-Shape number here includes PowerShell's COM marshalling**, which is
paid once per call and is not part of any backend. It is the same overhead in
both runs, so comparing the two backends is sound - but these are not the numbers
to quote for "how fast is an apply". They understate the accelerated backend's
advantage, because a large constant is added to both sides. The in-process
figures in benchmarks.md are the ones measured without a harness in the way.

Per-operation cost is the number that matters, and it is not the same as
throughput for the range apply: a range is one Office edit covering many Shapes,
so its per-Shape cost falls as the range grows while a per-Shape loop's does not.
That is the shape of the difference, and it is the same shape on both backends.

The skip cache is excluded from any "is portable fast" claim on purpose. A skip
costs nothing because it does nothing; quoting it as portable performance would
be describing an absent operation.

.PARAMETER Counts
Shape counts to measure. Defaults to 1, 32 and 100.

.PARAMETER Output
Where to write the report. Defaults to artifacts/backend_benchmarks.txt.
#>
param(
    [int[]]$Counts = @(1, 32, 100),
    [string]$Output
)

$ErrorActionPreference = 'Stop'

function Connect-PowerPoint {
    for ($attempt = 1; $attempt -le 15; $attempt++) {
        try {
            $candidate = New-Object -ComObject PowerPoint.Application
            $null = $candidate.Version
            $null = $candidate.Presentations.Count
            return $candidate
        } catch {
            if ($attempt -eq 15) { throw }
            Start-Sleep -Seconds 2
        }
    }
}

Add-Type -AssemblyName System.Drawing

$root = Split-Path $PSScriptRoot -Parent
if (-not $Output) { $Output = Join-Path $root 'artifacts/backend_benchmarks.txt' }
$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1

$rows = New-Object System.Collections.Generic.List[psobject]
function Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}

<#
Times one operation repeatedly and reports the distribution rather than a single
total. A median alongside the mean is worth having here because Office
occasionally stalls on something unrelated - an autosave, a redraw - and one
outlier moves a mean far more than it moves the truth.
#>
# $shapeCount, not $shapes: the scriptblocks below are invoked from inside this
# function and resolve their variables up the call stack, so a parameter sharing
# a name with the caller's Shape array would shadow it - and the block would see
# an integer where it expected the Shapes.
function Measure-Operation([string]$name, [int]$shapeCount, [int]$iterations, [scriptblock]$action) {
    $samples = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $iterations; $i++) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        & $action $i
        $watch.Stop()
        $samples.Add($watch.Elapsed.TotalMilliseconds)
    }
    $sorted = $samples | Sort-Object
    $median = if ($sorted.Count % 2 -eq 1) { $sorted[[int]($sorted.Count / 2)] }
              else { ($sorted[$sorted.Count / 2 - 1] + $sorted[$sorted.Count / 2]) / 2 }
    $total = ($samples | Measure-Object -Sum).Sum
    $mean = $total / $samples.Count

    $script:rows.Add([pscustomobject]@{
        Operation = $name
        Shapes = $shapeCount
        Calls = $iterations
        TotalMs = [Math]::Round($total, 2)
        MeanMs = [Math]::Round($mean, 3)
        MedianMs = [Math]::Round($median, 3)
        PerShapeMs = [Math]::Round($mean / [Math]::Max($shapeCount, 1), 4)
        PerSec = [Math]::Round(1000.0 / [Math]::Max($mean, 0.0001), 1)
    })
    Write-Host ("  {0,-44} {1,4} shapes  mean {2,8:N3} ms  median {3,8:N3} ms" -f
        $name, $shapeCount, $mean, $median)
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('bb_bench_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

function New-SolidPng([string]$path, [int]$r, [int]$g, [int]$b, [int]$size = 128) {
    $bitmap = New-Object System.Drawing.Bitmap($size, $size,
        [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, $r, $g, $b))
    $graphics.FillRectangle($brush, 0, 0, $size, $size)
    $graphics.Dispose(); $brush.Dispose()
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

$imageA = Join-Path $temp 'a.png'
$imageB = Join-Path $temp 'b.png'
$rewrite = Join-Path $temp 'rewrite.png'
New-SolidPng $imageA 220 40 40
New-SolidPng $imageB 40 80 220
New-SolidPng $rewrite 220 40 40

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$capabilities = $engine.AbiCapabilities()
$backend = Field $capabilities 'backend'
$mask = [uint32](Field $capabilities 'mask')
Write-Host "backend: $backend (mask 0x$('{0:X4}' -f $mask))"
Write-Host ''

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    [byte[]]$bytesA = [IO.File]::ReadAllBytes($imageA)
    [byte[]]$bytesB = [IO.File]::ReadAllBytes($imageB)
    $textureA = [long](Field ($engine.LoadTextureBytesAbi($bytesA)) 'texture')
    $textureB = [long](Field ($engine.LoadTextureBytesAbi($bytesB)) 'texture')

    foreach ($count in $Counts) {
        Write-Host "--- $count shape(s) ---"
        $names = @()
        $shapes = @()
        for ($i = 0; $i -lt $count; $i++) {
            $column = $i % 20
            $row = [int]($i / 20)
            $s = $slide.Shapes.AddShape($msoShapeRectangle, 10 + $column * 34, 10 + $row * 34, 30, 30)
            $s.Name = "Bench_${count}_$i"
            $names += $s.Name
            $shapes += $s
        }

        # Office's own API, called exactly as a caller would without BlipBridge.
        # This is the number everything else is worth comparing against.
        Measure-Operation "Fill.UserPicture (direct, per Shape)" $count 5 {
            param($iteration)
            $file = if ($iteration % 2 -eq 0) { $imageA } else { $imageB }
            foreach ($s in $shapes) { $null = $s.Fill.UserPicture($file) }
        }

        Measure-Operation "ApplyTexture (per Shape)" $count 5 {
            param($iteration)
            $handle = if ($iteration % 2 -eq 0) { $textureA } else { $textureB }
            foreach ($s in $shapes) { $null = $engine.ApplyTexture($s, $handle) }
        }

        # A miss every time: the texture alternates, so the cache can never skip.
        Measure-Operation "ApplyTextureIfChanged (miss)" $count 5 {
            param($iteration)
            $handle = if ($iteration % 2 -eq 0) { $textureA } else { $textureB }
            foreach ($s in $shapes) { $null = $engine.ApplyTextureIfChanged($s, $handle) }
        }

        # A hit every time. This measures the cost of deciding not to work, which
        # is a useful number and is not a measure of how fast an apply is.
        foreach ($s in $shapes) { $null = $engine.ApplyTextureIfChanged($s, $textureA) }
        Measure-Operation "ApplyTextureIfChanged (skip)" $count 5 {
            param($iteration)
            foreach ($s in $shapes) { $null = $engine.ApplyTextureIfChanged($s, $textureA) }
        }

        $range = $slide.Shapes.Range($names)
        Measure-Operation "ApplyTextureRange (one call)" $count 5 {
            param($iteration)
            $handle = if ($iteration % 2 -eq 0) { $textureA } else { $textureB }
            $null = $engine.ApplyTextureRange($range, $handle)
        }

        foreach ($s in $shapes) { $null = $s.Delete() }
        Write-Host ''
    }

    # --- ApplyPicture, whose three cases cost very different amounts ----------
    Write-Host '--- ApplyPicture ---'
    $pictureShape = $slide.Shapes.AddShape($msoShapeRectangle, 400, 300, 80, 80)
    $pictureShape.Name = 'BenchPicture'

    # Distinct files each time, so every call is a first sight of that image.
    $freshFiles = @()
    for ($i = 0; $i -lt 5; $i++) {
        $path = Join-Path $temp "fresh_$i.png"
        New-SolidPng $path (40 + $i * 30) 90 120
        $freshFiles += $path
    }
    Measure-Operation "ApplyPicture (first sight of the file)" 1 5 {
        param($iteration)
        $null = $engine.ApplyPicture($pictureShape, $freshFiles[$iteration])
    }

    $null = $engine.ApplyPicture($pictureShape, $imageA)
    Measure-Operation "ApplyPicture (cache hit, same file)" 1 5 {
        param($iteration)
        $null = $engine.ApplyPicture($pictureShape, $imageA)
    }

    # The correctness case: the same path, rewritten. It must cost a real apply,
    # because serving the cached image would be showing the wrong picture.
    Measure-Operation "ApplyPicture (same path, rewritten)" 1 5 {
        param($iteration)
        if ($iteration % 2 -eq 0) { New-SolidPng $rewrite 40 80 220 }
        else { New-SolidPng $rewrite 220 40 40 }
        $null = $engine.ApplyPicture($pictureShape, $rewrite)
    }
} finally {
    $presentation.Saved = $true
    $presentation.Close()
    $addin.Connect = $false
    Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
}

# Windows PowerShell has no ternary, and this script has to run under it.
if ([Environment]::Is64BitProcess) { $hostBitness = 'x64 host' } else { $hostBitness = 'x86 host' }

$environment = @(
    "backend      : $backend",
    "capabilities : 0x$('{0:X4}' -f $mask)",
    "PowerPoint   : $($app.Version) ($hostBitness)",
    "OS           : $([Environment]::OSVersion.VersionString)",
    "machine      : $([Environment]::ProcessorCount) logical processors",
    "measured     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    '',
    'Every per-Shape figure includes PowerShell COM marshalling, paid once per',
    'call and belonging to no backend. It is identical in both runs, so the two',
    'backends compare soundly; it is not the number to quote for how fast an',
    'apply is. See docs/benchmarks.md for in-process measurements.'
)

New-Item -ItemType Directory -Path (Split-Path $Output -Parent) -Force | Out-Null
$table = $rows | Format-Table Operation, Shapes, Calls, TotalMs, MeanMs, MedianMs, PerShapeMs, PerSec -AutoSize |
    Out-String -Width 160
($environment + '' + $table) | Set-Content $Output -Encoding utf8

Write-Host ''
Write-Host "report written to $Output"
$rows | Format-Table Operation, Shapes, MeanMs, MedianMs, PerShapeMs -AutoSize
