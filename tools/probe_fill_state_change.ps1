<#
.SYNOPSIS
Finds which object behind Shape.Fill holds the picture fill, by watching what
changes when Office applies one.

.DESCRIPTION
The strongest evidence available for "what is this object for" is what happens to
it when Office does the thing in question - and it can be had without calling
anything private. So: dump the chain, apply a picture fill through the
**documented** `Fill.UserPicture`, dump again, and compare.

An object whose digest changes when a picture fill is applied is holding some
part of that fill. An object whose digest does not change is not, whatever its
vtable is named after.

Everything here is read-only with respect to Office internals. The only thing
that writes is `Fill.UserPicture`, which is the public API, on a disposable
Shape in a disposable presentation.

.PARAMETER Chain
Offsets to follow from Shape.Fill, e.g. "4" or "4,0x20". Default "4".

.PARAMETER Output
Where to write the report.
#>
param([string]$Chain = '4', [string]$Output)

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
if (-not $Output) { $Output = Join-Path $root 'artifacts/fill_state_change_x86.txt' }
$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1

$temp = Join-Path ([IO.Path]::GetTempPath()) ('bb_state_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

# A tiny deterministic image: 2x2, four distinct corners. Small enough that
# nothing about it can be mistaken for scaling noise later.
$imagePath = Join-Path $temp 'probe.png'
$bitmap = New-Object System.Drawing.Bitmap(2, 2, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$bitmap.SetPixel(0, 0, [System.Drawing.Color]::FromArgb(255, 255, 0, 0))
$bitmap.SetPixel(1, 0, [System.Drawing.Color]::FromArgb(255, 0, 255, 0))
$bitmap.SetPixel(0, 1, [System.Drawing.Color]::FromArgb(255, 0, 0, 255))
$bitmap.SetPixel(1, 1, [System.Drawing.Color]::FromArgb(255, 255, 255, 255))
$bitmap.Save($imagePath, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

$second = Join-Path $temp 'probe2.png'
$bitmap = New-Object System.Drawing.Bitmap(2, 2, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
for ($y = 0; $y -lt 2; $y++) {
    for ($x = 0; $x -lt 2; $x++) {
        $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, 40, 90, 200))
    }
}
$bitmap.Save($second, [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Dispose()

$lines = New-Object System.Collections.Generic.List[string]
function Emit([string]$text) { $script:lines.Add($text); Write-Host $text }

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, 40, 40, 160, 120)

    Emit "chain followed : $Chain"
    Emit "image          : 2x2, four distinct corners"
    Emit ''

    <#
    Fill is re-fetched for every observation. PowerPoint hands back a wrapper
    each time it is asked, and holding one across an operation would risk
    observing a stale object rather than the current state - which would look
    exactly like "nothing changed".
    #>
    $before = $engine.DigestChain($shape.Fill, $Chain, 0x80)
    Emit "before any fill:"
    Emit "  $before"

    $null = $shape.Fill.UserPicture($imagePath)
    $afterFirst = $engine.DigestChain($shape.Fill, $Chain, 0x80)
    Emit ''
    Emit "after UserPicture (image A):"
    Emit "  $afterFirst"

    $null = $shape.Fill.UserPicture($second)
    $afterSecond = $engine.DigestChain($shape.Fill, $Chain, 0x80)
    Emit ''
    Emit "after UserPicture (image B):"
    Emit "  $afterSecond"

    $null = $shape.Fill.Solid()
    $afterSolid = $engine.DigestChain($shape.Fill, $Chain, 0x80)
    Emit ''
    Emit "after Fill.Solid (picture removed):"
    Emit "  $afterSolid"

    <#
    Reading the comparison. Each hop reports address, vtable and digest, so
    three different things can be told apart:

      - the object was replaced      (address changes)
      - the object was mutated       (address same, digest changes)
      - the object was untouched     (both same)

    Only the first two are objects that participate in a picture fill.
    #>
    Emit ''
    Emit 'comparison:'
    $names = @('before', 'imageA', 'imageB', 'solid')
    $samples = @($before, $afterFirst, $afterSecond, $afterSolid)
    $hops = ($before -split ';' | Where-Object { $_ }).Count
    for ($hop = 0; $hop -lt $hops; $hop++) {
        $row = @()
        for ($i = 0; $i -lt $samples.Count; $i++) {
            $field = ($samples[$i] -split ';' | Where-Object { $_ -like "hop$hop=*" })
            $row += $field
        }
        Emit "  hop $hop"
        for ($i = 0; $i -lt $row.Count; $i++) {
            Emit ("    {0,-8} {1}" -f $names[$i], $row[$i])
        }
    }

    # And the full structure of the last hop, with the picture fill in place, so
    # the fields that moved can be looked at rather than only counted.
    $null = $shape.Fill.UserPicture($imagePath)
    Emit ''
    Emit 'full dump with a picture fill applied:'
    Emit ($engine.InspectChain($shape.Fill, $Chain, 16, 0x80))
} finally {
    $presentation.Saved = $true
    $presentation.Close()
    $addin.Connect = $false
    Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
}

New-Item -ItemType Directory -Path (Split-Path $Output -Parent) -Force | Out-Null
$lines | Set-Content $Output -Encoding utf8
Write-Host ''
Write-Host "written to $Output"
