<#
.SYNOPSIS
Stability, Undo/Redo and retention checks for the native apply.

.DESCRIPTION
Covers the parts of the validation list that a single successful apply cannot:

  - Undo and Redo across a native apply must not crash and must leave the host
    usable.
  - Repeated applies must not destabilise PowerPoint or grow memory without
    bound, and closing the presentation must give that memory back.
  - Applying to a Shape whose presentation is then closed must not leave
    anything behind that a later ordinary operation trips over.

Memory is read from the host process, so the numbers include all of Office's own
document and undo state. They are a stability signal, not an allocation audit.
#>
param([int]$Iterations = 200)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_64_1.png'
$results = New-Object System.Collections.Generic.List[string]

$app = New-Object -ComObject PowerPoint.Application
$app.COMAddIns.Update()
$addin = $app.COMAddIns.Item('BlipBridge.Engine')
$addin.Connect = $true
$engine = $addin.Object
$host_ = Get-Process -Id $engine.GetHostProcessId()

function Get-PrivateMb { [math]::Round((Get-Process -Id $host_.Id).PrivateMemorySize64 / 1MB, 1) }

$presentation = $app.Presentations.Add(0)
try {
    $slide = $presentation.Slides.Add(1, 12)
    $warmup = $slide.Shapes.AddShape(1, 400, 300, 40, 40)
    $warmup.Fill.UserPicture($texture)
    $warmup.Delete()

    [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)

    # --- Undo and Redo around one native apply --------------------------------
    $undoTarget = $slide.Shapes.AddShape(1, 20, 20, 100, 100)
    $undoTarget.Name = 'BB_undo_target'
    $null = $engine.NativeApplyExperiment($undoTarget.Fill, $bytes)
    if ([int]$undoTarget.Fill.Type -ne 6) { throw 'Apply did not take before the Undo test' }

    try {
        $app.CommandBars.ExecuteMso('Undo')
        $results.Add("Undo after a native apply ran; Fill.Type is now $([int]$undoTarget.Fill.Type)")
        $app.CommandBars.ExecuteMso('Redo')
        $results.Add("Redo after a native apply ran; Fill.Type is now $([int]$undoTarget.Fill.Type)")
    } catch {
        $results.Add("Undo after a native apply was refused: $($_.Exception.Message.Split([Environment]::NewLine)[0])")
    }
    if ($host_.HasExited) { throw 'PowerPoint exited during Undo/Redo' }

    # Control: an ordinary UserPicture on a separate Shape, then Undo. If this
    # one works, the refusal above is because the native apply registers no undo
    # entry, not because Undo is broken.
    $control = $slide.Shapes.AddShape(1, 20, 150, 100, 100)
    $control.Fill.UserPicture($texture)
    $controlBefore = [int]$control.Fill.Type
    try {
        $app.CommandBars.ExecuteMso('Undo')
        $results.Add("Control: Undo after an ordinary UserPicture ran; Fill.Type $controlBefore -> $([int]$control.Fill.Type)")
    } catch {
        $results.Add("Control: Undo after an ordinary UserPicture was also refused: $($_.Exception.Message.Split([Environment]::NewLine)[0])")
    }
    if ($host_.HasExited) { throw 'PowerPoint exited during the Undo control' }
    $results.Add('Host survived Undo/Redo.')

    # --- Repeated applies -----------------------------------------------------
    $stress = $slide.Shapes.AddShape(1, 150, 20, 100, 100)
    $stress.Name = 'BB_stress_target'
    $identity = @($stress.Id, $stress.Name, $stress.Left, $stress.Top, $stress.Width,
                  $stress.Height, $stress.ZOrderPosition) -join '|'
    $before = Get-PrivateMb
    $watch = [Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $Iterations; $i++) {
        $null = $engine.NativeApplyExperiment($stress.Fill, $bytes)
    }
    $watch.Stop()
    $after = Get-PrivateMb
    $now = @($stress.Id, $stress.Name, $stress.Left, $stress.Top, $stress.Width,
             $stress.Height, $stress.ZOrderPosition) -join '|'
    if ($now -ne $identity) { throw 'Repeated applies changed Shape identity or geometry' }
    if ([int]$stress.Fill.Type -ne 6) { throw 'Repeated applies lost the picture fill' }
    $results.Add("$Iterations repeated native applies on one Shape: identity and picture fill intact.")
    $results.Add("Private bytes ${before} MB -> ${after} MB during the run (includes Office undo state).")
    $results.Add("Elapsed $([math]::Round($watch.Elapsed.TotalMilliseconds,1)) ms total; this is NOT a benchmark: each call re-decodes the image and rebuilds every record.")

    # An ordinary picture fill must still work afterwards.
    $check = $slide.Shapes.AddShape(1, 280, 20, 100, 100)
    $check.Fill.UserPicture($texture)
    if ([int]$check.Fill.Type -ne 6) { throw 'Ordinary UserPicture broke after the stress run' }
    $results.Add('Ordinary UserPicture still works after the stress run.')
} finally {
    $presentation.Saved = -1
    $presentation.Close()
}

Start-Sleep -Seconds 2
$results.Add("Private bytes after closing the presentation: $(Get-PrivateMb) MB")

# The host must still be able to do ordinary work in a fresh presentation.
$fresh = $app.Presentations.Add(0)
try {
    $shape = $fresh.Slides.Add(1, 12).Shapes.AddShape(1, 20, 20, 100, 100)
    $shape.Fill.UserPicture($texture)
    if ([int]$shape.Fill.Type -ne 6) { throw 'Host cannot fill Shapes in a fresh presentation' }
    $results.Add('Fresh presentation still works after everything above.')
} finally {
    $fresh.Saved = -1
    $fresh.Close()
}

$results | Set-Content "$root/artifacts/native_apply_stability.txt"
$results
