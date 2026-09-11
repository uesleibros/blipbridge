<#
.SYNOPSIS
Runs the Office runtime suites in order and reports one table.

.DESCRIPTION
Each suite drives a real PowerPoint, and most of them quit it when they finish.
Started back to back, the next suite can connect to a host that is still shutting
down - which surfaces as 0x800706B5 "unknown interface" from the COM activation,
or as a presentation that disappears underneath the run. Neither says anything
about BlipBridge, and either one turns a green matrix red for no reason.

So this waits for PowerPoint to be gone between suites rather than sleeping and
hoping, and reports every suite's own last line. A failure here is a failure of
the code under test, which is the only reason a release gate is worth having.

.EXAMPLE
.\tools\run_office_matrix.ps1
.EXAMPLE
.\tools\run_office_matrix.ps1 -Only test_range_apply, test_apply_if_changed
#>
param(
    [string[]]$Only = @(),
    [int]$SettleSeconds = 30
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

# In dependency-free order: the cheap contract checks first, so a broken build
# fails in seconds rather than after the stress suites.
$suites = @(
    'test_abi_in_powerpoint'
    'test_semantic_guards'
    'test_shape_compatibility'
    'test_native_apply'
    'test_native_texture'
    'test_picture_cache'
    'test_shape_lifecycle_cache'
    'test_cache_ownership'
    'test_undo_harness'
    'test_apply_if_changed'
    'test_range_apply'
    'test_image_api'
    'test_quad_warp'
    'test_native_texture_shutdown'
    'test_native_texture_stress'
    'test_release_stress'
)
if ($Only.Count -gt 0) { $suites = $Only }

function Wait-ForQuietHost([int]$seconds) {
    # A host with no presentation open is finishing with someone else's help;
    # one with a document open belongs to a person and is left alone.
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        $running = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
        if ($running.Count -eq 0) { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}

$rows = @()
foreach ($suite in $suites) {
    $script = Join-Path $root "tools/$suite.ps1"
    if (-not (Test-Path $script)) {
        $rows += [pscustomobject]@{ Suite = $suite; Result = 'MISSING'; Detail = '' }
        continue
    }

    Write-Host "running $suite ..."
    $started = Get-Date
    try {
        $output = & $script 2>&1 | Out-String
        $detail = ($output.Trim() -split "`r?`n" | Select-Object -Last 1).Trim()
        $rows += [pscustomobject]@{
            Suite   = $suite
            Result  = 'pass'
            Seconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
            Detail  = $detail
        }
    } catch {
        $rows += [pscustomobject]@{
            Suite   = $suite
            Result  = 'FAIL'
            Seconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
            Detail  = $_.Exception.Message.Split("`n")[0].Trim()
        }
    }

    [void](Wait-ForQuietHost $SettleSeconds)
}

$rows | Format-Table -AutoSize
$failed = @($rows | Where-Object { $_.Result -ne 'pass' })
if ($failed.Count -gt 0) {
    throw "$($failed.Count) suite(s) did not pass: $(($failed.Suite) -join ', ')"
}
"all $($rows.Count) Office suites passed"
