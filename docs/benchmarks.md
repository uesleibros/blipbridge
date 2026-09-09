# Benchmarks

## Native texture against Fill.UserPicture

Office 16.0.14334.20848 x64, measured in process by
`experiments/exp_internal_blip/texture_benchmark.cpp`, 1000 iterations on one
Shape with one image. Driving the loop from PowerShell adds a cross-process COM
round trip several times larger than the operation, so PowerShell timings
elsewhere in this repository are stability checks, not benchmarks.

| | mean | median | p95 | p99 | max |
|---|---:|---:|---:|---:|---:|
| `Fill.UserPicture(path)` | 0.6588 ms | 0.6189 ms | 0.8348 ms | 1.6002 ms | 2.6318 ms |
| `ApplyTexture`, same texture | 0.1860 ms | 0.1756 ms | 0.2260 ms | 0.3361 ms | 1.5537 ms |
| `ApplyTexture`, alternating two | 0.1857 ms | 0.1774 ms | 0.2309 ms | 0.2907 ms | 0.4037 ms |
| `LoadTexture` (50 samples) | 0.0231 ms | 0.0226 ms | 0.0235 ms | 0.0393 ms | 0.0393 ms |
| `ReleaseTexture` (50 samples) | 0.0001 ms | 0.0001 ms | 0.0001 ms | 0.0007 ms | 0.0007 ms |

Mean speed-up 3.54x, median 3.52x. Alternating between two textures costs the
same as repeating one, so switching carries no penalty. The tail is the larger
difference: p99 0.336 ms against 1.600 ms.

`LoadTexture` is paid once per texture and recovered after 0.05 of one apply.

The gap is real work, not a trick: Office re-decodes and writes a temporary image
per `UserPicture` call - 24 Content.MSO file events for 12 calls, against zero
for the native path. Full context in `native_texture.md`.

## Earlier fallback measurements

## Native texture against Fill.UserPicture

Office 16.0.14334.20848 x64, measured in process by
`experiments/exp_internal_blip/texture_benchmark.cpp`, 500 iterations on one
Shape with one image. Driving the loop from PowerShell adds a cross-process COM
round trip several times larger than the operation, so those numbers elsewhere in
the repository are stability checks, not benchmarks.

| | mean | median |
|---|---:|---:|
| `Fill.UserPicture(path)` | 0.6692 ms | 0.6337 ms |
| `ApplyTexture(handle)` | 0.1845 ms | 0.1706 ms |
| speed-up | 3.63x | 3.71x |

`LoadTexture` costs 0.0495 ms once per texture and is recovered after a tenth of
one apply.

The gap is real work, not a trick: Office re-decodes the image on every
`UserPicture` call, which a texture handle removes. Full context and the lifetime
matrix are in `native_texture.md`.

# Measured baseline

Office 16.0.14334.20848 x64; GCC/UCRT64 Release. Native code executes in POWERPNT.EXE on the owning STA, through the explicitly connected research COM add-in.

QPC around each operation; 10 warmup operations excluded. Each measurement includes late-bound GetIDsOfNames/Invoke and allocation overhead. Visible disposable presentation, no message pumping during the synchronous batch. These are operation-completion timings, not measured screen refresh FPS. Run order is fixed; undo history and growing document state are confounders.

| Operation | N | Total ms | Mean us | Median us | Min us | Max us | Ops/s |
|---|---|---|---|---|---|---|---|
| UserPicture_same_64 | 100 | 111.992 | 1119.92 | 1054.7 | 753.5 | 1931.9 | 892.919 |
| UserPicture_alternating_64 | 100 | 142.091 | 1420.91 | 1279.1 | 756.5 | 4801.4 | 703.772 |
| Apply_prePicked | 100 | 70.925 | 709.25 | 663.3 | 598.4 | 1690.3 | 1409.94 |
| PickUp_Apply_alternating | 100 | 110.801 | 1108.01 | 1022.5 | 763.9 | 2285.3 | 902.516 |
| UserPicture_same_64 | 1000 | 1211.9 | 1211.9 | 1153.6 | 738.7 | 2933.9 | 825.148 |
| UserPicture_alternating_64 | 1000 | 1255.62 | 1255.62 | 1194.1 | 742 | 4221.5 | 796.421 |
| Apply_prePicked | 1000 | 901.555 | 901.555 | 816.4 | 628.3 | 3251.4 | 1109.2 |
| PickUp_Apply_alternating | 1000 | 1112.22 | 1112.22 | 994.9 | 794.1 | 4728.6 | 899.105 |
| UserPicture_same_64 | 10000 | 13766 | 1376.6 | 1266.6 | 760.7 | 9036.5 | 726.429 |
| UserPicture_alternating_64 | 10000 | 16264.5 | 1626.45 | 1539.1 | 882.4 | 12250.3 | 614.838 |
| Apply_prePicked | 10000 | 14121.6 | 1412.16 | 1312.2 | 762.6 | 6789.5 | 708.136 |
| PickUp_Apply_alternating | 10000 | 19095.8 | 1909.58 | 1806.9 | 1043.8 | 9777.1 | 523.674 |
| UserPicture_same_32 | 1000 | 2355.06 | 2355.06 | 2358.7 | 1365.3 | 5812 | 424.618 |
| UserPicture_same_128 | 1000 | 2491.71 | 2491.71 | 2475.9 | 1324.3 | 6594.1 | 401.331 |
| UserPicture_same_256 | 1000 | 2845.4 | 2845.4 | 2751.6 | 1647.2 | 8276.8 | 351.444 |
| Shape_creation | 1000 | 781.722 | 781.722 | 617.1 | 400.6 | 4154.5 | 1279.23 |
| Shape_deletion | 1000 | 898.921 | 898.921 | 748.7 | 458.4 | 3528.5 | 1112.44 |
| Duplicate_and_delete | 1000 | 1860.42 | 1860.42 | 1684.9 | 929.2 | 7589.6 | 537.512 |
| Node_SetPosition | 1000 | 5173.72 | 5173.72 | 5167.1 | 744.7 | 17378.9 | 193.285 |
| Visibility_change | 1000 | 794.17 | 794.17 | 628.7 | 428 | 3977 | 1259.18 |
| Frame130_UserPicture | 100 | 127800 | 1.278e+06 | 1.27886e+06 | 884631 | 1.75411e+06 | 0.782475 |
| Frame130_PickUp_Apply | 100 | 214581 | 2.14581e+06 | 2.21748e+06 | 1.73747e+06 | 2.58967e+06 | 0.466025 |

Frame130 rows: one operation is a frame of 130 existing freeforms, alternating between two textures across shapes and frames, one node edit and one visibility update per shape. No creation/deletion occurs in the measured frame. PickUp/Apply copies other formatting too.

MemoryStream and CachedBlip: not implemented; no timings claimed. VBA wrapper measurements: module provided but not yet executed. External-process results are retained separately and are not a valid in-process performance comparison.

Raw CSVs are copied to benchmarks/results. External run ended with RPC_E_CALL_REJECTED during shape churn. The in-process run blocks the UI while executing, and memory growth includes Office undo/document state; this is not a proven native leak.

A debugger attach/detach paused PowerPoint for approximately three seconds during the visible fallback frame run. Together with fixed ordering/history, this prevents treating the visible comparison as a controlled final speedup measurement.

## focused

| Operation | N | Total ms | Mean us | Median us | Min us | Max us | Ops/s |
|---|---|---|---|---|---|---|---|
| Hidden_Invoke_UserPicture | 1000 | 1257.12 | 1257.12 | 1089.1 | 734.9 | 30862.9 | 795.466 |
| Hidden_Dual_UserPicture_same | 1000 | 1090.29 | 1090.29 | 998.9 | 719.9 | 4930.8 | 917.187 |
| Hidden_Dual_UserPicture_alternate | 1000 | 1069.58 | 1069.58 | 995.3 | 724.6 | 4245.2 | 934.948 |
| Hidden_Dual_Apply | 10000 | 8668.11 | 866.811 | 775.9 | 577.8 | 5253.9 | 1153.65 |
| Hidden_Dual_PickUp_Apply | 10000 | 12503.4 | 1250.34 | 1138.4 | 784.5 | 8691.2 | 799.781 |
| Hidden_Invoke_PickUp_Apply | 1000 | 1334.8 | 1334.8 | 1287.4 | 874.4 | 3129.1 | 749.175 |

## stress

| Operation | N | Total ms | Mean us | Median us | Min us | Max us | Ops/s |
|---|---|---|---|---|---|---|---|
| MemoryAdapter_including_hooks_logging | 100 | 243.711 | 2437.11 | 2303.3 | 1885.5 | 4563.5 | 410.322 |
| CachedDonor_100000_alternating | 100000 | 244203 | 2442.03 | 2394 | 710.1 | 20387.6 | 409.495 |
| CachedDonor_130_individual_fills | 20 | 10191.4 | 509568 | 506170 | 486853 | 534349 | 1.96245 |
| CachedDonor_130_ShapeRange_fills | 100 | 174809 | 1.74809e+06 | 1.73993e+06 | 1.56089e+06 | 2.10983e+06 | 0.572054 |

Focused tests use a hidden disposable presentation and public dual-COM methods resolved from live type information. They do not measure the VBA wrapper.

Stress ran 100,000 alternating donor applications successfully. Private bytes grew from about 171 MB to 1.01 GB, then to 3.04 GB after subsequent individual/range batches, returning to 225 MB after document close. The saved deck reopened with 133 normal picture-filled shapes, no Picture shapes and two media resources. These sequential history-heavy workloads do not establish a leak or controlled range-batching speedup.

MemoryAdapter timing includes experimental hooks/logging and predates WIC input validation. Office writes a temporary Content.MSO PNG during this path; this is not a compliant memory-only backend timing.
