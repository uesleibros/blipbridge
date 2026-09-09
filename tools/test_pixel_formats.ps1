<#
.SYNOPSIS
Finds which ARC::SurfaceFormat values the raw-pixel creator accepts, and what
channel order each one means.

.DESCRIPTION
GFX exports a cached-image creator taking a raw pixel buffer, a stride and an
ARC::SurfaceFormat. The enum has no symbols, so this determines it empirically
rather than guessing: build a buffer of one known colour, create a texture with
each candidate format value, apply it to a Shape, export the Shape and read the
rendered colour back.

The source colour is deliberately asymmetric - R=0xF0, G=0x40, B=0x10 - so the
rendered pixel identifies the channel order unambiguously. A format that swaps
red and blue is obvious; a symmetric colour like grey would hide it.

Nothing is assumed about which values are valid. A value that throws is reported
as rejected, which is itself the useful answer.
#>
param([int]$MaxFormat = 24, [int]$Size = 16)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$results = New-Object System.Collections.Generic.List[string]
Add-Type -AssemblyName System.Drawing

$msoTrue = -1
$ppLayoutBlank = 12
$msoShapeRectangle = 1
$msoFillPicture = 6

# Byte order in the buffer is B, G, R, A - the ordinary Windows 32bpp layout.
# If a format value interprets it differently, the exported colour will say so.
$srcR = 0xF0; $srcG = 0x40; $srcB = 0x10
$stride = $Size * 4
[byte[]]$pixels = New-Object byte[] ($stride * $Size)
for ($y = 0; $y -lt $Size; $y++) {
    for ($x = 0; $x -lt $Size; $x++) {
        $o = $y * $stride + $x * 4
        $pixels[$o + 0] = $srcB
        $pixels[$o + 1] = $srcG
        $pixels[$o + 2] = $srcR
        $pixels[$o + 3] = 0xFF
    }
}

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$presentation = $app.Presentations.Add(0)
$export = Join-Path $root 'artifacts/pixel_format_probe.png'
try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    # Warm GFX/OART with an ordinary fill on a throwaway Shape.
    $warm = $slide.Shapes.AddShape($msoShapeRectangle, 500, 350, 40, 40)
    $warm.Fill.UserPicture((Join-Path $root 'artifacts/textures/texture_64_1.png'))
    $warm.Delete()

    $results.Add("source colour R=$('{0:X2}' -f $srcR) G=$('{0:X2}' -f $srcG) B=$('{0:X2}' -f $srcB), buffer order BGRA, ${Size}x${Size}, stride $stride")

    for ($format = 0; $format -le $MaxFormat; $format++) {
        $shape = $slide.Shapes.AddShape($msoShapeRectangle, 40, 40, 120, 120)
        try {
            $report = $engine.PixelTextureExperiment($shape.Fill, $pixels, $Size, $Size, $stride, $format)
            if ([int]$shape.Fill.Type -ne $msoFillPicture) {
                $results.Add("format $format : created but produced no picture fill")
            } else {
                $shape.Export($export, 2)
                $bitmap = [System.Drawing.Bitmap]::FromFile($export)
                try {
                    $c = $bitmap.GetPixel([int]($bitmap.Width / 2), [int]($bitmap.Height / 2))
                } finally { $bitmap.Dispose() }
                $match = if ($c.R -eq $srcR -and $c.G -eq $srcG -and $c.B -eq $srcB) { 'BGRA-in-memory MATCH' }
                         elseif ($c.R -eq $srcB -and $c.G -eq $srcG -and $c.B -eq $srcR) { 'channels swapped (RGBA order)' }
                         else { 'other' }
                $results.Add("format $format : rendered R=$('{0:X2}' -f $c.R) G=$('{0:X2}' -f $c.G) B=$('{0:X2}' -f $c.B) A=$('{0:X2}' -f $c.A) -> $match")
            }
        } catch {
            $line = $_.Exception.Message.Split([Environment]::NewLine)[0]
            $results.Add("format $format : rejected - $($line.Split('|')[0].Trim())")
        } finally {
            $shape.Delete()
        }
    }
    # --- prove the argument mapping ------------------------------------------
    # Every format value rendering correctly could mean the format is tolerated,
    # or it could mean the value never reaches the format slot. Distinguish the
    # two by varying arguments whose effect is not in doubt: a two-colour image
    # renders as two halves only if width and stride land where intended, and a
    # deliberately wrong stride must visibly skew it.
    $half = [int]($Size / 2)
    [byte[]]$halves = New-Object byte[] ($stride * $Size)
    for ($y = 0; $y -lt $Size; $y++) {
        for ($x = 0; $x -lt $Size; $x++) {
            $o = $y * $stride + $x * 4
            if ($x -lt $half) { $halves[$o+0] = 0x10; $halves[$o+1] = 0x40; $halves[$o+2] = 0xF0 }
            else              { $halves[$o+0] = 0xF0; $halves[$o+1] = 0x40; $halves[$o+2] = 0x10 }
            $halves[$o+3] = 0xFF
        }
    }

    function Get-RenderedHalves([byte[]]$buffer, [int]$w, [int]$h, [int]$s, [int]$fmt) {
        $shape = $slide.Shapes.AddShape($msoShapeRectangle, 40, 40, 160, 160)
        try {
            $null = $engine.PixelTextureExperiment($shape.Fill, $buffer, $w, $h, $s, $fmt)
            $shape.Export($export, 2)
            $bitmap = [System.Drawing.Bitmap]::FromFile($export)
            try {
                $l = $bitmap.GetPixel([int]($bitmap.Width * 0.25), [int]($bitmap.Height / 2))
                $r = $bitmap.GetPixel([int]($bitmap.Width * 0.75), [int]($bitmap.Height / 2))
                return "L=$('{0:X2}{1:X2}{2:X2}' -f $l.R,$l.G,$l.B) R=$('{0:X2}{1:X2}{2:X2}' -f $r.R,$r.G,$r.B)"
            } finally { $bitmap.Dispose() }
        } finally { $shape.Delete() }
    }

    $correct = Get-RenderedHalves $halves $Size $Size $stride 0
    $results.Add("mapping/correct stride  : $correct  (expect L=F04010 R=1040F0)")
    try {
        $skewed = Get-RenderedHalves $halves $Size ($Size - 1) ($stride + 4) 0
        $results.Add("mapping/stride +4       : $skewed  (must differ if stride is honoured)")
        if ($skewed -eq $correct) {
            $results.Add('  WARNING: stride had no effect - the argument mapping is NOT confirmed')
        } else {
            $results.Add('  stride changes the render, so width/stride land where intended.')
        }
    } catch {
        $results.Add("mapping/stride +4       : rejected - $($_.Exception.Message.Split([Environment]::NewLine)[0])")
    }
} finally {
    $presentation.Saved = $msoTrue
    $presentation.Close()
}

$results | Set-Content "$root/artifacts/pixel_formats.txt"
$results
