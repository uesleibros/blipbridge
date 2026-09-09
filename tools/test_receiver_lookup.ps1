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

    # Object-kind matrix: the structural guard must accept every genuine Fill
    # wrapper and reject everything else with a diagnosis rather than a walk.
    $range = $first.Shapes.Range(1)
    $kinds = [ordered]@{
        'Shape.Fill'      = $shapeA.Fill
        'ShapeRange.Fill' = $range.Fill
        'Shape'           = $shapeA
        'Shape.Line'      = $shapeA.Line
        'Shape.TextFrame' = $shapeA.TextFrame
        'Slide'           = $first
    }
    foreach ($kind in $kinds.Keys) {
        try {
            $null = $engine.InspectFillReceiver($kinds[$kind])
            $results.Add("$kind -> accepted")
        } catch {
            $first_line = $_.Exception.Message.Split([Environment]::NewLine)[0]
            $results.Add("$kind -> rejected: $($first_line.Split('|')[0].Trim())")
        }
    }
    foreach ($mustWork in 'Shape.Fill', 'ShapeRange.Fill') {
        $null = $engine.InspectFillReceiver($kinds[$mustWork])
    }
    foreach ($mustFail in 'Shape', 'Shape.Line', 'Shape.TextFrame', 'Slide') {
        $rejected = $false
        try { $null = $engine.InspectFillReceiver($kinds[$mustFail]) } catch { $rejected = $true }
        if (-not $rejected) { throw "$mustFail was accepted but is not a FillFormat" }
    }
    $results.Add('Both Fill wrappers accepted; Shape, Line, TextFrame and Slide all rejected.')

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
