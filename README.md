# BlipBridge

Experimental C++20 native research project for assigning and reusing image fills on existing PowerPoint AutoShapes and Freeforms from VBA. Built and tested locally with MSYS2 UCRT64 GCC on Windows x64.

**PNG/JPEG Byte arrays can reach existing fills on PowerPoint LTSC 2021 build 16.0.14334.20848, but Office creates a temporary PNG internally. The strict no-temporary-image milestone and an internal decoded-BLIP cache remain unsolved.**

The normal COM backend is explicitly `PickupApplyFallback`: it retains a preloaded donor Shape and transfers its style to existing targets. PickUp/Apply copies line and other formatting as well as the fill. It is suitable only when that behavior is acceptable.

The separate `MemoryFillExperiment` temporarily adapts Office source-file API imports to a memory buffer. BlipBridge supplies bytes without an input image file, but the write trace proves Office copies them to `INetCache\Content.MSO\<name>.png` and then decodes that cache through an IStream adapter. This disqualifies the experiment as the requested no-temporary-image solution. It is build/signature guarded, explicitly invoked research instrumentation, and is not used by normal LoadTexture/SetImageBytes. Those methods currently return an unsupported error.

## Verified results

- Detected Office version, bitness, toolchain and dynamically loaded modules.
- Ran in-process QPC benchmarks with 100, 1,000 and 10,000 fill operations and a pooled 130-freeform workload.
- Traced UserPicture from PPCORE through OART and MSO20 to the full compressed-image read, preserving real RVAs and call stacks.
- Fed 32/64/128/256 PNGs and JPEG from Byte arrays into existing shapes using the experimental memory adapter, with the Office-internal temporary-file limitation above.
- Verified freeform identity and geometry, editable nodes, saving and reopening. The final PNG export is pixel-identical to ordinary UserPicture before and after reopen. Its saved media bytes match the source SHA256.
- Saved baseline: 134 normal image-filled shapes, zero Picture shapes, two shared media resources. Serialized deduplication does not establish decoded-memory reuse.

See [status](docs/status.md), [native pipeline](docs/userpicture_pipeline.md), [research journal](docs/research.md), [environment](docs/environment.md) and [measurements](docs/benchmarks.md).

## Build and registration

From 64-bit PowerShell:

```powershell
.\tools\office_probe.ps1 -StartPowerPoint
.\tools\generate_textures.ps1
.\build.ps1 -Configuration Release
.\tools\register.ps1
```

Debug: `.\build.ps1 -Configuration Debug`. The probe first discovers existing tools, including the installed MSYS2 UCRT64 environment. No MSVC dependency or System32 copying is required. Registration is per-user x64; unregister with `.\tools\unregister.ps1`. Close PowerPoint before rebuilding a loaded DLL; Windows locks loaded modules.

From UCRT64 shell:

```bash
cmake -S . -B build/ucrt64 -G 'MinGW Makefiles' -DCMAKE_BUILD_TYPE=Release
cmake --build build/ucrt64 --parallel
```

## VBA fallback

Import the modules under `vba/`, or call Automation directly:

```vb
Dim bb As Object, texture As Long
Set bb = CreateObject("BlipBridge.Engine")
Debug.Print bb.GetBackendName()
texture = bb.RegisterTextureShape(ActivePresentation.Slides(1).Shapes("texture_donor"))
bb.ApplyTexture ActivePresentation.Slides(1).Shapes("poly_17"), texture
bb.ReleaseTexture texture
```

The donor must already be a normal Shape with a picture fill. Keep its presentation open while using the handle. ClearTextures before closing the document. Handles own COM references, are local to an Engine instance, and are not native pointers. Calls execute synchronously on the owning STA. VBA wrappers are supplied; separate VBA interpreter/wrapper timings have not yet been run.

Creating the Engine from VBA loads it in POWERPNT.EXE. Creating it directly from PowerShell loads it in PowerShell instead. For native in-process experiments, explicitly connect the research COM add-in:

```powershell
.\tools\register.ps1 -ResearchAddin
.\tools\run_inproc.ps1
.\tools\trace_userpicture.ps1
.\tools\test_memory.ps1
.\tools\run_stress.ps1
.\tests\com_smoke.ps1
```

The add-in is registered with LoadBehavior=0 and connected by these scripts. They create disposable presentations. Long synchronous batches block the PowerPoint UI and accumulate Office undo/document memory; do not use them during interactive editing. Scripts do not change Office security settings.

## Performance and limitations

The initial visible-document in-process 10,000-call averages and all distributions are in [the benchmark report](docs/benchmarks.md). A separate hidden-document dual-COM experiment measured about 1.09 ms for UserPicture and 0.87 ms for a pre-picked Apply; PickUp plus Apply was about 1.25 ms. These are measured operation completion times, not screen refresh FPS. The public fallback is not universally faster.

The memory adapter still runs Office's ordinary image materialization pipeline, including its own temporary image copy. A negative test showed Office can report success for invalid bytes, so the current experiment first validates PNG/JPEG through Windows WIC. This adds a validation decode before Office's work. Limits: 64 MiB compressed and 16 megapixels decoded. Profile checks, hook setup and logging add research overhead. No compressed source data is processed by the public donor-application code, but internal decoder invocations during Apply have not yet been comprehensively counted. Internal stream ABI, decoded-resource ownership, fill-only binding, broad compatibility and production reliability remain open work.

License: MIT for this project. No Office binaries or PDBs are distributed.
