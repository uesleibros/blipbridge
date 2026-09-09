<#
.SYNOPSIS
Reproducible driver for a read-only GDB observation of live PowerPoint.

.DESCRIPTION
Starts a preparation script in a background PowerShell job, waits for it to
publish artifacts/decoder_target.json, then attaches GDB with the requested
probe. Every supported probe only reads memory: it makes no inferior call and
detaches before the preparation job finishes.

The preparation script is expected to publish artifacts/decoder_target.json and
then block until artifacts/decoder_go appears. This driver writes that trigger
and always releases the job, including when the probe aborts early.

.EXAMPLE
.\tools\run_office_probe.ps1
Runs the fill-transaction observation with its default transcript path.

.EXAMPLE
.\tools\run_office_probe.ps1 -Prepare tools/prepare_receiver_identity.ps1 `
    -Probe experiments/exp_internal_blip/probe_receiver_identity.py `
    -Output artifacts/receiver_identity.txt
#>
param(
    [string]$Prepare = 'tools/prepare_decoder_trace.ps1',
    [string]$Probe = 'experiments/exp_internal_blip/probe_fill_transaction.py',
    [string]$Output = 'artifacts/fill_transaction_lifecycle.txt',
    [int]$ReadyTimeoutSeconds = 90
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

# Stale trigger or target files would let the probe race a previous run.
Remove-Item -LiteralPath "$root/artifacts/decoder_go" -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "$root/artifacts/decoder_target.json" -ErrorAction SilentlyContinue

$environment = Get-Content "$root/artifacts/environment.json" -Raw | ConvertFrom-Json
$gdb = Join-Path $environment.bin 'gdb.exe'
if (!(Test-Path -LiteralPath $gdb)) { throw "GDB not found at $gdb" }

$job = Start-Job -ScriptBlock {
    param($script)
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script
} -ArgumentList (Join-Path $root $Prepare)

try {
    $deadline = (Get-Date).AddSeconds($ReadyTimeoutSeconds)
    while (!(Test-Path -LiteralPath "$root/artifacts/decoder_target.json")) {
        if ((Get-Date) -gt $deadline) { throw 'Trace preparation did not publish a target' }
        if ($job.State -eq 'Failed') { throw 'Trace preparation failed' }
        Start-Sleep -Milliseconds 200
    }

    Push-Location $root
    try {
        # GDB reports every Office worker thread; those lines carry no
        # evidence and would dominate the archived transcript.
        & $gdb -batch -x $Probe 2>&1 |
            Where-Object { $_ -notmatch '^\[(New|Thread) ' } |
            Set-Content -LiteralPath (Join-Path $root $Output) -Encoding utf8
    } finally {
        Pop-Location
    }
} finally {
    # Release the preparation script even when the probe aborted early.
    if (!(Test-Path -LiteralPath "$root/artifacts/decoder_go")) {
        Set-Content -LiteralPath "$root/artifacts/decoder_go" -Value 'go'
    }
    Wait-Job $job -Timeout 120 | Out-Null
    Receive-Job $job -ErrorAction Continue
    Remove-Job $job -Force
    Remove-Item -LiteralPath "$root/artifacts/decoder_go" -ErrorAction SilentlyContinue
}

"Transcript: $Output"
