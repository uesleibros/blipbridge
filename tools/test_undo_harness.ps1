<#
.SYNOPSIS
Undo/Redo harness, control case first.

.DESCRIPTION
Earlier attempts concluded nothing because the control failed too:
CommandBars.ExecuteMso('Undo') returned E_FAIL even after an ordinary
Fill.UserPicture. The likely cause was the harness, not the fill - those runs
used Presentations.Add(0), which creates a presentation with **no window**, and
PowerPoint's Undo acts on a document window.

So this script does three things in order:

  1. Creates a visible, windowed presentation.
  2. Finds an undo driver that demonstrably works on an ordinary
     Fill.UserPicture. Several are tried; the first that actually reverts the
     fill is adopted.
  3. Runs the identical sequence against the native ApplyTexture.

Only step 3's result is a statement about the native path, and only if step 2
succeeded. If no driver can undo a normal UserPicture, the run reports the
harness as unusable rather than blaming the native path.
#>
param([switch]$KeepOpen)

$ErrorActionPreference = 'Stop'

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

$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

$msoTrue = -1
$msoFalse = 0
$ppLayoutBlank = 12
$msoShapeRectangle = 1
$msoFillSolid = 1
$msoFillPicture = 6
$undoCommandBarId = 128   # the classic Undo control, still present under the ribbon
$redoCommandBarId = 129


$app = Connect-PowerPoint
$app.Visible = $msoTrue
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object

# A windowed presentation. This is the part the earlier harness got wrong.
$presentation = $app.Presentations.Add($msoTrue)

<#
 Each driver returns $true if it believes it issued an Undo. Whether the Undo
 actually happened is decided by the caller from the document, never from the
 driver's own return value.
#>
$undoDrivers = [ordered]@{
    'ExecuteMso' = {
        param($application)
        $application.CommandBars.ExecuteMso('Undo')
        return $true
    }
    'CommandBarControl' = {
        param($application)
        $control = $application.CommandBars.FindControl($null, $undoCommandBarId, $null, $null)
        if (-not $control) { return $false }
        $control.Execute()
        return $true
    }
    'SendKeys' = {
        param($application)
        $application.Activate()
        Start-Sleep -Milliseconds 250
        [void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [System.Windows.Forms.SendKeys]::SendWait('^z')
        Start-Sleep -Milliseconds 400
        return $true
    }
}
$redoDrivers = [ordered]@{
    'ExecuteMso' = {
        param($application)
        $application.CommandBars.ExecuteMso('Redo')
        return $true
    }
    'CommandBarControl' = {
        param($application)
        $control = $application.CommandBars.FindControl($null, $redoCommandBarId, $null, $null)
        if (-not $control) { return $false }
        $control.Execute()
        return $true
    }
    'SendKeys' = {
        param($application)
        $application.Activate()
        Start-Sleep -Milliseconds 250
        [void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [System.Windows.Forms.SendKeys]::SendWait('^y')
        Start-Sleep -Milliseconds 400
        return $true
    }
}

function New-TargetShape($slide, $left) {
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, $left, 40, 120, 120)
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = 0x3366CC
    return $shape
}

try {
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    # Warm GFX/OART on a throwaway Shape so module loading is never inside a
    # measured or observed sequence.
    $warmup = $slide.Shapes.AddShape($msoShapeRectangle, 500, 300, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()
    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)

    # ---------------------------------------------------------------- control
    # Find a driver that actually undoes an ordinary UserPicture.
    $workingDriver = $null
    foreach ($name in $undoDrivers.Keys) {
        $probe = New-TargetShape $slide 40
        $app.StartNewUndoEntry()
        $probe.Fill.UserPicture($texture)
        if ([int]$probe.Fill.Type -ne $msoFillPicture) { throw 'Control fill did not apply' }

        $issued = $false
        $failure = $null
        try { $issued = & $undoDrivers[$name] $app } catch { $failure = $_.Exception.Message.Split([Environment]::NewLine)[0] }

        $reverted = $false
        try { $reverted = [int]$probe.Fill.Type -ne $msoFillPicture } catch { $reverted = $true }

        if ($failure) {
            $results.Add("control/$name : driver raised - $failure")
        } elseif ($reverted) {
            $results.Add("control/$name : UNDO WORKS (UserPicture reverted)")
            $workingDriver = $name
        } else {
            $results.Add("control/$name : driver ran but the fill did not revert")
        }
        try { $probe.Delete() } catch { }
        if ($workingDriver) { break }
    }

    if (-not $workingDriver) {
        $results.Add('HARNESS UNUSABLE: no driver can undo an ordinary Fill.UserPicture.')
        $results.Add('Nothing can be concluded about native undo from this run.')
    } else {
        $results.Add("Adopted undo driver: $workingDriver")

        # Confirm Redo works too, on the control, before judging the native path.
        $probe = New-TargetShape $slide 180
        $app.StartNewUndoEntry()
        $probe.Fill.UserPicture($texture)
        & $undoDrivers[$workingDriver] $app | Out-Null
        $undone = [int]$probe.Fill.Type -ne $msoFillPicture
        $redone = $false
        try {
            & $redoDrivers[$workingDriver] $app | Out-Null
            $redone = [int]$probe.Fill.Type -eq $msoFillPicture
        } catch { }
        $results.Add("control/redo : undo=$undone redo=$redone")
        try { $probe.Delete() } catch { }

        # ------------------------------------------------------------- native
        $handle = $engine.LoadTexture($bytes)
        try {
            $target = New-TargetShape $slide 320
            $before = [int]$target.Fill.Type
            $app.StartNewUndoEntry()
            $null = $engine.ApplyTexture($target, $handle)
            $applied = [int]$target.Fill.Type
            if ($applied -ne $msoFillPicture) { throw 'Native apply did not produce a picture fill' }

            $failure = $null
            try { & $undoDrivers[$workingDriver] $app | Out-Null }
            catch { $failure = $_.Exception.Message.Split([Environment]::NewLine)[0] }
            $afterUndo = -1
            try { $afterUndo = [int]$target.Fill.Type } catch { $afterUndo = -2 }

            if ($failure) {
                $results.Add("native/undo : driver raised - $failure")
            } else {
                $results.Add("native/undo : Fill.Type $before -> $applied -> $afterUndo (undo reverted: $($afterUndo -ne $msoFillPicture))")
            }

            $afterRedo = -1
            try {
                & $redoDrivers[$workingDriver] $app | Out-Null
                $afterRedo = [int]$target.Fill.Type
            } catch { $afterRedo = -2 }
            $results.Add("native/redo : Fill.Type after redo = $afterRedo (restored: $($afterRedo -eq $msoFillPicture))")

            # Whatever undo did, the host and the Shape must still be sane.
            if ($app.Presentations.Count -lt 1) { throw 'PowerPoint lost the presentation during undo' }
            $probe2 = New-TargetShape $slide 460
            $probe2.Fill.UserPicture($texture)
            if ([int]$probe2.Fill.Type -ne $msoFillPicture) { throw 'Host broken after native undo' }
            $results.Add('Host and ordinary UserPicture still healthy after the native undo sequence.')
        } finally {
            $engine.ReleaseTexture($handle)
        }
    }
} finally {
    if (-not $KeepOpen) {
        $presentation.Saved = $msoTrue
        $presentation.Close()
        $app.Quit()
    }
}

$results | Set-Content "$root/artifacts/undo_harness.txt"
$results
