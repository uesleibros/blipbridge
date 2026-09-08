$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$rows=Import-Csv "$root/artifacts/baseline_inproc.csv"
$out=@('# Measured baseline','', 'Office 16.0.14334.20848 x64; GCC/UCRT64 Release. Native code executes in POWERPNT.EXE on the owning STA, through the explicitly connected research COM add-in.', '', 'QPC around each operation; 10 warmup operations excluded. Each measurement includes late-bound GetIDsOfNames/Invoke and allocation overhead. Visible disposable presentation, no message pumping during the synchronous batch. These are operation-completion timings, not measured screen refresh FPS. Run order is fixed; undo history and growing document state are confounders.', '', '| Operation | N | Total ms | Mean us | Median us | Min us | Max us | Ops/s |','|---|---|---|---|---|---|---|---|')
foreach($r in $rows){$out+="| $($r.operation) | $($r.count) | $($r.total_ms) | $($r.average_us) | $($r.median_us) | $($r.min_us) | $($r.max_us) | $($r.ops_per_sec) |"}
$out+=@('', 'Frame130 rows: one operation is a frame of 130 existing freeforms, alternating between two textures across shapes and frames, one node edit and one visibility update per shape. No creation/deletion occurs in the measured frame. PickUp/Apply copies other formatting too.', '', 'MemoryStream and CachedBlip: not implemented; no timings claimed. VBA wrapper measurements: module provided but not yet executed. External-process results are retained separately and are not a valid in-process performance comparison.', '', 'Raw CSVs are copied to benchmarks/results. External run ended with RPC_E_CALL_REJECTED during shape churn. The in-process run blocks the UI while executing, and memory growth includes Office undo/document state; this is not a proven native leak.')
$out+=@('', 'A debugger attach/detach paused PowerPoint for approximately three seconds during the visible fallback frame run. Together with fixed ordering/history, this prevents treating the visible comparison as a controlled final speedup measurement.')
foreach($dataset in 'focused','stress'){
 $file="$root/artifacts/$dataset.csv"
 if(Test-Path -LiteralPath $file){
  $out+=@('',"## $dataset",'', '| Operation | N | Total ms | Mean us | Median us | Min us | Max us | Ops/s |','|---|---|---|---|---|---|---|---|')
  foreach($r in (Import-Csv $file | Where-Object {$_.operation -notlike '#*'})){
   $out+="| $($r.operation) | $($r.count) | $($r.total_ms) | $($r.average_us) | $($r.median_us) | $($r.min_us) | $($r.max_us) | $($r.ops_per_sec) |"
  }
 }
}
$out+=@('', 'Focused tests use a hidden disposable presentation and public dual-COM methods resolved from live type information. They do not measure the VBA wrapper.', '', 'Stress ran 100,000 alternating donor applications successfully. Private bytes grew from about 171 MB to 1.01 GB, then to 3.04 GB after subsequent individual/range batches, returning to 225 MB after document close. The saved deck reopened with 133 normal picture-filled shapes, no Picture shapes and two media resources. These sequential history-heavy workloads do not establish a leak or controlled range-batching speedup.', '', 'MemoryAdapter timing includes experimental hooks/logging and predates WIC input validation. Office writes a temporary Content.MSO PNG during this path; this is not a compliant memory-only backend timing.')
$out | Set-Content "$root/docs/benchmarks.md"
New-Item -ItemType Directory -Force "$root/benchmarks/results" | Out-Null
Copy-Item "$root/artifacts/baseline*.csv" "$root/benchmarks/results" -Force
