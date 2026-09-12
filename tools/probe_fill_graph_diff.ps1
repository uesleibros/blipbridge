<#
.SYNOPSIS
Finds the object that holds a Shape's picture fill, by diffing the reachable
object graph across a documented fill.

.DESCRIPTION
Guessing pointer chains by hand stops working after the second hop. This walks
everything reachable from `Shape.Fill` that looks like an object, takes that
snapshot either side of an operation, and reports what differed.

Three outcomes are told apart, and they mean different things:

  appeared / disappeared   an object was created or released by the operation
  mutated                  same address, contents changed - it holds state
  unchanged                not involved

The operation is `Fill.UserPicture`, the **documented** API. Nothing private is
called; the whole point is to learn where the fill lands before anything is
called at all.

Two fills with *different images* are compared as well as fill against no-fill.
An object that changes between two different pictures is holding something
picture-specific, which is a much narrower claim than "changes when a fill
happens" - a dirty flag would do the latter.

.PARAMETER Depth
How far to walk. Default 3.

.PARAMETER Nodes
Maximum objects to collect. Default 800.
#>
param([int]$Depth = 3, [int]$Nodes = 800, [string]$Output)

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
if (-not $Output) { $Output = Join-Path $root 'artifacts/fill_graph_diff_x86.txt' }
$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1

$temp = Join-Path ([IO.Path]::GetTempPath()) ('bb_graph_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null

function New-Solid([string]$path, [int]$r, [int]$g, [int]$b, [int]$size) {
    $bitmap = New-Object System.Drawing.Bitmap($size, $size,
        [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    for ($y = 0; $y -lt $size; $y++) {
        for ($x = 0; $x -lt $size; $x++) {
            $bitmap.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, $r, $g, $b))
        }
    }
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()
}

# Two images that differ in colour *and* in size. A size difference is the more
# useful of the two: a field holding width or height will show up as a changed
# scalar somewhere, which is far easier to recognise than a changed pointer.
$imageA = Join-Path $temp 'a.png'
$imageB = Join-Path $temp 'b.png'
New-Solid $imageA 220 30 30 8
New-Solid $imageB 30 30 220 64

$lines = New-Object System.Collections.Generic.List[string]
function Emit([string]$text) { $script:lines.Add($text); Write-Host $text }

function Parse([string]$snapshot) {
    $map = @{}
    foreach ($line in ($snapshot -split "`r?`n")) {
        if ($line -notmatch '^0x') { continue }
        $parts = $line -split '\|'
        if ($parts.Count -lt 5) { continue }
        $map[$parts[0]] = [pscustomobject]@{
            Address = $parts[0]; Vtable = $parts[1]; Digest = $parts[2]
            Depth = $parts[3]; Path = $parts[4]
        }
    }
    return $map
}

function Compare-Snapshots($left, $right, [string]$label) {
    Emit ''
    Emit "=== $label ==="
    $mutated = @()
    foreach ($address in $left.Keys) {
        if (-not $right.ContainsKey($address)) { continue }
        if ($left[$address].Digest -ne $right[$address].Digest) { $mutated += $right[$address] }
    }
    $appeared = @($right.Keys | Where-Object { -not $left.ContainsKey($_) })
    $gone = @($left.Keys | Where-Object { -not $right.ContainsKey($_) })

    Emit ("  objects before {0}, after {1}" -f $left.Count, $right.Count)
    Emit ("  mutated {0}, appeared {1}, disappeared {2}" -f $mutated.Count, $appeared.Count, $gone.Count)
    if ($mutated.Count -gt 0) {
        Emit '  mutated objects:'
        foreach ($node in ($mutated | Sort-Object Depth)) {
            Emit ("    {0}  {1}  {2}  {3}" -f $node.Address, $node.Vtable, $node.Depth, $node.Path)
        }
    }
    if ($appeared.Count -gt 0 -and $appeared.Count -le 25) {
        Emit '  objects that appeared:'
        foreach ($address in $appeared) {
            $node = $right[$address]
            Emit ("    {0}  {1}  {2}  {3}" -f $node.Address, $node.Vtable, $node.Depth, $node.Path)
        }
    }
}

$app = Connect-PowerPoint
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

$presentation = $app.Presentations.Add($msoTrue)
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, 40, 40, 160, 120)

    Emit "walk depth $Depth, up to $Nodes objects, from Shape.Fill"
    Emit "image A: 8x8 red     image B: 64x64 blue"

    # Touch the Fill once first, so lazy initialisation is not mistaken for the
    # effect of the fill. That confusion cost a whole earlier reading.
    $null = $shape.Fill.Type
    $baseline = Parse $engine.DigestGraph($shape.Fill, $Depth, $Nodes, 0x80)

    $null = $shape.Fill.UserPicture($imageA)
    $withA = Parse $engine.DigestGraph($shape.Fill, $Depth, $Nodes, 0x80)

    $null = $shape.Fill.UserPicture($imageB)
    $withB = Parse $engine.DigestGraph($shape.Fill, $Depth, $Nodes, 0x80)

    $null = $shape.Fill.Solid()
    $solid = Parse $engine.DigestGraph($shape.Fill, $Depth, $Nodes, 0x80)

    Compare-Snapshots $baseline $withA 'no fill  ->  picture A'
    Compare-Snapshots $withA $withB 'picture A  ->  picture B (different image AND size)'
    Compare-Snapshots $withB $solid 'picture B  ->  solid colour'
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
