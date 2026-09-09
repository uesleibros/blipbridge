<#
.SYNOPSIS
Checks whether the OART receiver is reachable from a public FillFormat.

.DESCRIPTION
Calls the read-only InspectFillReceiver research method for several Shapes and
compares the reported fields. This repeats, without a debugger, the
classification that tools/prepare_receiver_identity.ps1 established under GDB:
the receiver must be stable per Shape and its container stable per slide.

Nothing here calls a private Office function. The research method only reads
memory behind vtable identity checks and reports borrowed addresses.

The Shapes are also inspected before and after an ordinary Fill.UserPicture, so
the report shows whether the receiver survives a fill unchanged.
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
$results = New-Object System.Collections.Generic.List[string]
try {
    $first = $presentation.Slides.Add(1, 12)
    $second = $presentation.Slides.Add(2, 12)
    $shapeA = $first.Shapes.AddShape(1, 10, 10, 100, 100)
    $shapeB = $first.Shapes.AddShape(1, 140, 10, 100, 100)
    $shapeC = $second.Shapes.AddShape(1, 10, 10, 100, 100)

    function Get-Field([string]$report, [string]$name) {
        foreach ($part in $report.Split(';')) {
            $pair = $part.Split('=', 2)
            if ($pair.Length -eq 2 -and $pair[0] -eq $name) { return $pair[1] }
        }
        return $null
    }

    # Each call re-evaluates $shape.Fill, so a stable result also proves the
    # receiver is not an artifact of one FillFormat wrapper instance.
    $before = @{
        A1 = $engine.InspectFillReceiver($shapeA.Fill)
        A2 = $engine.InspectFillReceiver($shapeA.Fill)
        B  = $engine.InspectFillReceiver($shapeB.Fill)
        C  = $engine.InspectFillReceiver($shapeC.Fill)
    }
    foreach ($key in 'A1', 'A2', 'B', 'C') {
        $results.Add("$key before fill: $($before[$key])")
    }

    if ((Get-Field $before.A1 'receiver') -ne (Get-Field $before.A2 'receiver')) {
        throw 'Receiver is not stable across two lookups on the same Shape'
    }
    if ((Get-Field $before.A1 'receiver') -eq (Get-Field $before.B 'receiver')) {
        throw 'Two Shapes on one slide reported the same receiver'
    }
    if ((Get-Field $before.A1 'container') -ne (Get-Field $before.B 'container')) {
        throw 'Shapes on one slide reported different containers'
    }
    if ((Get-Field $before.A1 'container') -eq (Get-Field $before.C 'container')) {
        throw 'Shapes on different slides reported the same container'
    }
    $results.Add('Receiver is stable per Shape; container is stable per slide.')

    $shapeA.Fill.UserPicture($texture)
    $afterA = $engine.InspectFillReceiver($shapeA.Fill)
    $results.Add("A after fill:    $afterA")
    if ((Get-Field $before.A1 'receiver') -ne (Get-Field $afterA 'receiver')) {
        throw 'Receiver changed across an ordinary UserPicture'
    }
    $results.Add('Receiver survives an ordinary UserPicture unchanged.')

    # A Shape deleted through public COM must not leave a usable lookup.
    $doomed = $first.Shapes.AddShape(1, 270, 10, 60, 60)
    $doomedFill = $doomed.Fill
    $doomed.Delete()
    try {
        $null = $engine.InspectFillReceiver($doomedFill)
        $results.Add('Deleted Shape still reported a receiver; treat lookups as unsafe after Delete.')
    } catch {
        $results.Add('Deleted Shape lookup failed as expected.')
    }
} finally {
    $results | Set-Content "$root/artifacts/receiver_lookup.txt"
    $results
    $presentation.Saved = -1
    $presentation.Close()
}
