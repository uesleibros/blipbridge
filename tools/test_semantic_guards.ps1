<#
.SYNOPSIS
Permanent regression suite for the semantic Shape guard.

.DESCRIPTION
Structural validation proves the internal objects are the Office objects we
expect. Semantic validation proves the operation is meaningful for that class of
Shape. A Connector passes the first and fails the second, and applying to one
terminated PowerPoint - so these cases stay in the suite for good.

Each case runs in its own PowerPoint process, via tools/semantic_guard_one.ps1,
and each asserts something a "did it crash?" test cannot:

  Connector  refused, and the private OART apply was never entered
  Line       refused, and the private OART apply was never entered
  WordArt    applied natively, still WordArt, text and geometry intact,
             persists through save/reopen, undoes and redoes

The Connector and Line cases check the entry counter the backend keeps, so a
refusal that happened *after* the dangerous call would fail here even if the
host survived it.

Output: artifacts/semantic_guards.txt and a printed table. Non-zero exit on any
failure, so this can gate a release.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$texture = Join-Path $root 'artifacts/textures/texture_128_0.png'
if (-not (Test-Path $texture)) { throw "Missing $texture - run tools/generate_textures.ps1" }

function Get-Field([string]$report, [string]$name) {
    foreach ($pair in $report -split ';') {
        $bits = $pair -split '=', 2
        if ($bits.Length -eq 2 -and $bits[0] -eq $name) { return $bits[1] }
    }
    return $null
}

$child = Join-Path $PSScriptRoot 'semantic_guard_one.ps1'
$rows = New-Object System.Collections.Generic.List[object]
$raw = New-Object System.Collections.Generic.List[string]

foreach ($case in @('Connector', 'Line', 'WordArt')) {
    Write-Host "semantic guard: $case ..." -NoNewline
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    $outFile = Join-Path $env:TEMP ('bb_guard_' + [guid]::NewGuid().ToString('N') + '.txt')
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $child),
        '-Case', ('"{0}"' -f $case), '-Out', ('"{0}"' -f $outFile)
    )
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments `
        -PassThru -WindowStyle Hidden
    if (-not $process.WaitForExit(300000)) { try { $process.Kill() } catch { } }

    $line = ''
    if (Test-Path $outFile) { $line = (Get-Content $outFile -Raw).Trim() }
    Remove-Item $outFile -ErrorAction SilentlyContinue

    if (-not $line) {
        $rows.Add([pscustomobject]@{
            Case = $case; Result = 'Crashed'; Eligibility = '-'
            Entries = '-'; Failed = 'child wrote no row'
        })
        $raw.Add("$case|Crashed|child wrote no row")
        Write-Host ' Crashed'
        continue
    }
    $raw.Add($line)

    $entriesBefore = Get-Field $line 'entriesBefore'
    $entriesAfter = Get-Field $line 'entriesAfter'
    $rows.Add([pscustomobject]@{
        Case        = $case
        Result      = Get-Field $line 'result'
        Eligibility = Get-Field $line 'eligibility'
        ShapeType   = Get-Field $line 'shapeType'
        Entries     = "$entriesBefore -> $entriesAfter"
        Healthy     = Get-Field $line 'hostHealthy'
        Failed      = Get-Field $line 'failed'
    })
    Write-Host (' {0}' -f (Get-Field $line 'result'))
}

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# BlipBridge semantic guard regressions')
foreach ($line in $raw) { $lines.Add($line) }
$lines | Set-Content "$root/artifacts/semantic_guards.txt"

''
$rows | Format-Table -AutoSize Case, Result, Eligibility, ShapeType, Entries, Healthy
$rows | Where-Object { $_.Failed } | ForEach-Object { "  $($_.Case): $($_.Failed)" }

$bad = @($rows | Where-Object { $_.Result -ne 'Passed' })
if ($bad.Count -gt 0) {
    throw "$($bad.Count) semantic guard case(s) did not pass"
}
'Semantic guards: Connector and Line refused before the private apply; WordArt verified.'
