<#
.SYNOPSIS
Isolated reproduction: does a native apply to a Connector kill PowerPoint?

.DESCRIPTION
The compatibility matrix run died at the Connector row with RPC 0x800706BE - the
host process gone mid-call. That is a serious claim, so it is reproduced here on
its own, with nothing else in the document and every step journalled before it is
attempted.

Three cases, each in a fresh presentation:

  A  control     an AutoShape, native apply           expected: survives
  B  connector   a Connector, ordinary UserPicture    expected: survives or refuses
  C  connector   a Connector, native apply            the case under test

If C kills the host while A and B do not, the finding is real and specific.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$msoTrue = -1
$ppLayoutBlank = 12
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'

$journal = "$root/artifacts/connector_isolation.txt"
Set-Content $journal "connector isolation $(Get-Date -Format o)"
function Note([string]$text) { Add-Content $journal $text; $text }

function Test-Case([string]$label, [scriptblock]$makeShape, [bool]$native) {
    $app = New-Object -ComObject PowerPoint.Application
    $app.COMAddIns.Update()
    $addin = $app.COMAddIns.Item('BlipBridge.Engine')
    $addin.Connect = $true
    $engine = $addin.Object
    $presentation = $app.Presentations.Add($msoTrue)
    $pid0 = $engine.GetHostProcessId()
    Note "$label : host pid $pid0"
    try {
        $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
        $shape = & $makeShape $slide
        Note "$label : created, Shape.Type=$($shape.Type) Name=$($shape.Name)"
        try {
            $probe = $engine.ProbeShapeCompatibility($shape)
            Note "$label : probe $probe"
        } catch { Note "$label : probe threw - $($_.Exception.Message -replace '\s+', ' ')" }

        Note "$label : about to apply (native=$native)"
        try {
            if ($native) {
                [byte[]]$bytes = [IO.File]::ReadAllBytes($texture)
                $handle = $engine.LoadTexture($bytes)
                $null = $engine.ApplyTexture($shape, $handle)
                $engine.ReleaseTexture($handle)
            } else {
                $shape.Fill.UserPicture($texture)
            }
            Note "$label : apply returned"
        } catch {
            Note "$label : apply threw - $($_.Exception.Message -replace '\s+', ' ')"
        }

        # The decisive question: is the host still there afterwards?
        $alive = $false
        try { $null = $app.Version; $alive = $true } catch { }
        Note "$label : host alive after apply = $alive"
        if ($alive) {
            try { Note "$label : Fill.Type=$($shape.Fill.Type) Shape.Type=$($shape.Type)" } catch {
                Note "$label : shape unreadable afterwards"
            }
        }
    } finally {
        try { $presentation.Saved = $msoTrue; $presentation.Close() } catch { }
        try { $app.Quit() } catch { }
        Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
    }
}

Test-Case 'A control autoshape native' { param($slide) $slide.Shapes.AddShape(1, 20, 20, 120, 120) } $true
Test-Case 'B connector userpicture'    { param($slide) $slide.Shapes.AddConnector(1, 60, 60, 200, 200) } $false
Test-Case 'C connector native'         { param($slide) $slide.Shapes.AddConnector(1, 60, 60, 200, 200) } $true

''
'===== summary ====='
Get-Content $journal
