# CI and releases

Two workflows.

## `ci.yml`

Runs on every push to `main` and every pull request. Four jobs: Windows **x64**
and **x86**, each in Release and Debug. Each one configures, builds, confirms the
produced DLL really is the architecture it claims, runs CTest, and checks that
every export the VBA wrapper binds to by name is present and undecorated.

The architecture check is not ceremony. A toolchain misconfiguration that
silently produced an x64 DLL in the x86 job would make every other check in that
job meaningless. The export check matters most on x86, where a lost
`-Wl,--kill-at` would decorate the stdcall exports as `BB_Init@0` and no VBA
`Declare` would find them.

## `release.yml`

Runs on a `v*` tag. Rebuilds from scratch, re-runs the tests, packages, generates
SHA-256 checksums, and publishes a GitHub Release. It does not publish if any
build or test job fails.

Only architectures with a **validated** native backend are packaged as supported
releases; see the release workflow for how x86 is labelled today.

## What CI does not do

**It does not prove the native PowerPoint backend works.**

The Office integration suites in `tools/` need a live PowerPoint on a validated
Office build, which hosted runners do not have. They are not run here and are
never reported as passing. They run on a machine with Office 16.0.14334.20848,
and their transcripts are committed under `docs/evidence/` so the evidence
travels with the repository even though the automation cannot reproduce it.

| Suite | Where | Needs Office |
|---|---|---|
| `abi_contract` | CI (both architectures) and locally | no - asserts fail-closed behaviour |
| `com_contract` | CI (x64) and locally | no |
| `tools/test_*.ps1` | validated machine only | yes |

A green tick is a **build** signal. Backend validation is tracked separately in
the README, and the two must never be conflated.

## Running the Office suites locally

On a machine with the validated Office build:

```powershell
.\build.ps1 -Configuration Release
.\tools\test_semantic_guards.ps1        # Connector, Line, WordArt
.\tools\test_shape_compatibility.ps1    # the full matrix, one process per class
.\tools\test_shape_class_stress.ps1     # 1000 applies per native class
.\tools\test_picture_cache.ps1          # UserPicture2 and both caches
.\tools\test_shape_lifecycle_cache.ps1  # duplicate, group, delete, undo
.\tools\test_native_texture.ps1         # lifetime matrix
.\tools\test_undo_harness.ps1           # undo and redo
.\tools\test_apply_if_changed.ps1       # the skip: what it skips, and what it must not
.\tools\test_range_apply.ps1            # the multi-Shape apply, and its undo behaviour
.\tools\test_image_api.ps1              # the CPU image resource, crop/orient/scale, quad warp
.\tools\test_quad_warp.ps1              # the projective warp, read back out of a render
.\tools\check_vba_module.ps1            # the wrapper's compile rules, without a VBA compiler
```

Or run the whole Office matrix in one go, which is what a release gate wants:

```powershell
.\tools\run_office_matrix.ps1
```

It waits for PowerPoint to be gone between suites rather than sleeping and
hoping. Started back to back, a suite can otherwise connect to a host the
previous one is still shutting down, which fails with 0x800706B5 "unknown
interface" and says nothing about BlipBridge.

Each writes a transcript under `artifacts/`; `tools/archive_evidence.ps1` copies
the ones worth keeping into `docs/evidence/`.

## A future self-hosted runner

`ci.yml` is written so an Office-capable runner can be added without
restructuring it: add a job that targets a self-hosted label and calls the
scripts above. Until such a runner exists, the Office suites stay out of CI
rather than being stubbed, skipped, or reported as passing.
