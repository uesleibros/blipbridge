<#
.SYNOPSIS
Stresses every Shape class that is native or a candidate for it, one process each.

.DESCRIPTION
The compatibility matrix says which classes survive an apply. This says whether
they survive a thousand of them, plus undo, save, reopen, deletion and close -
which is the bar a class has to clear before the allowlist is widened to include
it.

Runs tools/stress_one_class.ps1 per class in its own PowerPoint, for the same
reason the matrix does: a class that takes the host down must cost one row, not
the whole run.

Output: artifacts/shape_class_stress.txt and a printed table.
#>
param([int]$Applies = 1000, [int]$Alternating = 500)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

# Already allowlisted, plus every NativeCandidate the matrix produced. Connector
# and Line are deliberately absent: their fatality is established and re-running
# it proves nothing.
$classes = @(
    'AutoShape', 'Freeform', 'WordArt', 'Group child',
    'TextBox', 'Placeholder', 'Callout', 'Group', 'Picture', 'Media'
)

function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}

$child = Join-Path $PSScriptRoot 'stress_one_class.ps1'
$rows = New-Object System.Collections.Generic.List[object]

foreach ($class in $classes) {
    Write-Host "stressing $class ..." -NoNewline
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    $outFile = Join-Path $env:TEMP ('bb_stress_' + [guid]::NewGuid().ToString('N') + '.txt')
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $child),
        '-Category', ('"{0}"' -f $class), '-Out', ('"{0}"' -f $outFile),
        '-Applies', $Applies, '-Alternating', $Alternating
    )
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments `
        -PassThru -WindowStyle Hidden
    if (-not $process.WaitForExit(600000)) { try { $process.Kill() } catch { } }

    $line = ''
    if (Test-Path $outFile) { $line = (Get-Content $outFile -Raw).Trim() }
    Remove-Item $outFile -ErrorAction SilentlyContinue

    if (-not $line) {
        $rows.Add([pscustomobject]@{ Class = $class; Result = 'Crashed'; Stage = 'no row' })
        Write-Host ' Crashed (no row)'
        continue
    }

    # Composed first: a -f expression spanning lines does not parse inside a
    # hashtable literal in Windows PowerShell.
    $refsApplies = Get-Field $line 'refsAfterApplies'
    $refsAlternating = Get-Field $line 'refsAfterAlternating'
    $refsClose = Get-Field $line 'refsAfterClose'
    $refs = "$refsApplies/$refsAlternating/$refsClose"
    $privateStart = Get-Field $line 'privateStart'
    $privateEnd = Get-Field $line 'privateEnd'
    $private = "$privateStart->$privateEnd MB"

    $rows.Add([pscustomobject]@{
        Class    = $class
        Result   = Get-Field $line 'result'
        Stage    = Get-Field $line 'stage'
        Path     = Get-Field $line 'path'
        MsApply  = Get-Field $line 'msPerApply'
        Refs     = $refs
        Undo     = Get-Field $line 'undoRedo'
        Reopen   = Get-Field $line 'reopen'
        Delete   = Get-Field $line 'deleteThenApply'
        Private  = $private
        Handles  = Get-Field $line 'handlesAtEnd'
        Healthy  = Get-Field $line 'hostHealthy'
        Raw      = $line
    })
    Write-Host (' {0} ({1})' -f (Get-Field $line 'result'), (Get-Field $line 'stage'))
}

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# BlipBridge Shape class stress')
$lines.Add("# applies=$Applies alternating=$Alternating")
foreach ($row in $rows) { $lines.Add($row.Raw) }
$lines | Set-Content "$root/artifacts/shape_class_stress.txt"

''
$rows | Format-Table -AutoSize Class, Result, Path, MsApply, Refs, Undo, Reopen, Delete, Handles, Healthy
'Refs column is after-applies / after-alternating / after-close.'
'Full detail in artifacts/shape_class_stress.txt'
