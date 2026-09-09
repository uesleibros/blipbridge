<#
.SYNOPSIS
Builds the Shape compatibility matrix, one category per PowerPoint process.

.DESCRIPTION
The native backend accepts AutoShape and Freeform because those are the classes
whose receiver chain was proved. This finds out what the other classes actually
present, rather than inferring it from the fact that they expose a `Fill`.

Each category runs in its **own** PowerPoint process, via tools/shape_probe_one.ps1.
That is not caution for its own sake: a Connector reports msoAutoShape, presents
the identical wrapper, FillFormat and receiver as a rectangle, and a native apply
to one terminated PowerPoint. A single-process matrix loses every row after such
a class. Here the parent just sees a child that never wrote its row, and records
`Crashed` - which is itself the most important thing this matrix can report.

Classifications:

  NativeSupported    applied natively, identity/geometry/type intact, verified
  FallbackSupported  native refused, but ordinary Fill.UserPicture works
  Unsupported        neither path works, or verification failed
  CrashedDuringResearch  the host did not survive the test
  NotApplicable      no instance could be created on this machine

Output: artifacts/shape_matrix.txt (machine-readable) and a printed table.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'
if (-not (Test-Path $texture)) { throw "Missing $texture - run tools/generate_textures.ps1" }

# msoShapeType, for reading a category's reported Type back in the table.
$shapeTypeNames = @{
    1 = 'msoAutoShape'; 2 = 'msoCallout'; 3 = 'msoChart'; 5 = 'msoFreeform'; 6 = 'msoGroup'
    7 = 'msoEmbeddedOLEObject'; 9 = 'msoLine'; 11 = 'msoLinkedPicture'; 13 = 'msoPicture'
    14 = 'msoPlaceholder'; 15 = 'msoTextEffect'; 16 = 'msoMedia'; 17 = 'msoTextBox'
    19 = 'msoTable'; 21 = 'msoDiagram'; 24 = 'msoIgraphic'; 26 = 'msoSmartArt'
}

$categories = @(
    'AutoShape', 'Freeform', 'TextBox', 'Placeholder', 'Callout',
    'Connector', 'Line', 'Group', 'Group child',
    'Picture', 'WordArt', 'Table', 'Chart', 'SmartArt', 'OLE', 'Media'
)

function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}

$child = Join-Path $PSScriptRoot 'shape_probe_one.ps1'
$results = New-Object System.Collections.Generic.List[object]

foreach ($category in $categories) {
    Write-Host "testing $category ..." -NoNewline
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    $outFile = Join-Path $env:TEMP ('bb_row_' + [guid]::NewGuid().ToString('N') + '.txt')
    # Every value is quoted: Start-Process joins an ArgumentList with spaces and
    # quotes nothing, so "Group child" would otherwise arrive as two arguments.
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $child),
        '-Category', ('"{0}"' -f $category), '-Out', ('"{0}"' -f $outFile)
    )
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments `
        -PassThru -WindowStyle Hidden
    # Generous: a Chart or SmartArt can take a while to materialise the first time.
    if (-not $process.WaitForExit(180000)) {
        try { $process.Kill() } catch { }
    }

    $line = ''
    if (Test-Path $outFile) { $line = (Get-Content $outFile -Raw).Trim() }
    Remove-Item $outFile -ErrorAction SilentlyContinue

    if (-not $line) {
        # The child wrote nothing at all, which it only does if it died before
        # its very first write, or could not start.
        $results.Add([pscustomobject]@{
            Category = $category; Class = 'CrashedDuringResearch'; Type = '-'
            Structural = '-'; Semantic = '-'; Route = '-'; Fill = '-'
            Undo = '-'; Redo = '-'; Reopen = '-'; Survived = 'no'
            Detail = "child exited $($process.ExitCode) with no row"
        })
        Write-Host ' CrashedDuringResearch (no row)'
        continue
    }

    $typeValue = Get-Field $line 'shapeType'
    $typeName = if ($typeValue -and $shapeTypeNames.ContainsKey([int]$typeValue)) {
        "$typeValue/$($shapeTypeNames[[int]$typeValue])"
    } elseif ($typeValue) { "$typeValue" } else { '-' }

    $class = Get-Field $line 'class'
    $detailParts = New-Object System.Collections.Generic.List[string]
    foreach ($key in 'detail', 'unrestricted', 'fallback', 'undo', 'reopen', 'connector', 'inner') {
        $value = Get-Field $line $key
        if ($value) { $detailParts.Add("$key=$value") }
    }

    $fillBefore = Get-Field $line 'fillBefore'
    $fillAfter = Get-Field $line 'fillType'
    $results.Add([pscustomobject]@{
        Category   = $category
        Class      = $class
        Type       = $typeName
        Structural = Get-Field $line 'structural'
        Semantic   = Get-Field $line 'semantic'
        Route      = Get-Field $line 'route'
        Fill       = "$fillBefore->$fillAfter"
        Undo       = Get-Field $line 'undo'
        Redo       = Get-Field $line 'redo'
        Reopen     = Get-Field $line 'reopen'
        Survived   = 'yes'
        Detail     = ($detailParts -join '; ')
    })
    Write-Host " $class"
}

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force

# The build is a property of the machine, recorded once rather than per row.
$environment = Get-Content (Join-Path $root 'artifacts/environment.json') -Raw | ConvertFrom-Json
$build = $environment.officeVersion
if (-not $build) { $build = 'unknown' }

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# BlipBridge Shape compatibility matrix')
$lines.Add("# build=$build")
$lines.Add('# category|class|shapeType|structural|semantic|route|fill|undo|redo|reopen|survived|detail')
foreach ($row in $results) {
    $lines.Add(('{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}' -f `
        $row.Category, $row.Class, $row.Type, $row.Structural, $row.Semantic, $row.Route,
        $row.Fill, $row.Undo, $row.Redo, $row.Reopen, $row.Survived, $row.Detail))
}
$lines | Set-Content "$root/artifacts/shape_matrix.txt"

''
"PowerPoint build: $build"
$results | Format-Table -AutoSize Category, Class, Type, Structural, Semantic, Route, Fill, Undo, Redo, Survived
'Full detail in artifacts/shape_matrix.txt'
