# BlipBridge

Fast, reusable image textures for ordinary PowerPoint Shapes, callable from VBA.

```vb
BlipBridge.UserPicture2 shp, "C:\textures\brick.png"   ' that is the whole API
```

It decides the rest: the accelerated path for Shape classes that have one,
Office's own `Fill.UserPicture` for classes that do not, the file read and
decoded once, and no work at all when that Shape already carries that image. For
full control the texture handles are still there:

```vb
tex = BlipBridge.LoadTexture(bytes)      ' decode once
BlipBridge.ApplyTexture shp, tex         ' ~0.19 ms, no file, no donor Shape
BlipBridge.ReleaseTexture tex
```

No `regsvr32`. No ProgID. No `CreateObject`. No add-in installer. Put the DLL
next to your presentation and import one `.bas` module.

> **Status: research-grade, working, and narrow.** The accelerated backend runs
> against **one validated Office build** and refuses to run against anything
> else. Read [Supported builds](#supported-builds) before depending on it.

## Why it exists

`Fill.UserPicture` is the only supported way to put an image on a Shape's fill
from VBA, and it re-does all of the work every time. Measured on the test
machine, each call writes a temporary PNG under `INetCache\Content.MSO`, decodes
it again, and builds a fresh internal image object - even when you pass the same
file to the same Shape twice in a row.

For a document that repaints many Shapes per frame, that is the whole budget.

BlipBridge decodes an image **once** into the same internal cached-image object
PowerPoint uses, and then applies that object to as many Shapes as you like.

## What it does, in one paragraph

`LoadTexture` wraps your bytes in an in-memory stream and calls PowerPoint's own
exported cached-image creator, giving back a handle. `ApplyTexture` builds the
same property record PowerPoint's own `UserPicture` handler builds, puts the
already-decoded image into it, and commits it through the same internal
transaction. The Shape is not replaced, moved, restyled or converted - only its
fill changes. Save, reopen, Undo and Redo all behave exactly as they do for a
normal picture fill.

## Performance

Measured in process on the test machine, 1000 iterations, same Shape and image:

| | mean | median | p95 | p99 |
|---|---:|---:|---:|---:|
| `Fill.UserPicture(path)` | 0.6588 ms | 0.6189 ms | 0.8348 ms | 1.6002 ms |
| `BlipBridge.ApplyTexture` | **0.1860 ms** | **0.1756 ms** | **0.2260 ms** | **0.3361 ms** |

**≈3.5x faster on the mean, ≈3.5x on the median**, and a wider margin in the
tail. One-off costs: `LoadTexture` 0.0231 ms, `ReleaseTexture` 0.0001 ms.
Alternating between two textures costs the same as repeating one.

These numbers are from one Office build, one machine and one workload. Your
mileage will differ with Office version, hardware, image size and how much else
the document is doing. Re-measure with `tools/run_texture_benchmark.ps1` rather
than trusting the table.

**Raw pixels are available but not faster.** `LoadTexturePixels` takes a BGRA32
buffer and skips image decoding, which matters if you already hold pixels and
would otherwise encode a PNG first. Against an *already encoded* PNG of the same
size it is ~15-18% slower. See [docs/pixel_textures.md](docs/pixel_textures.md).

**Changing content costs more, but less than first reported.** Images cannot be
mutated in place - no Office export writes to one, and the cached image is
content-addressed by a unique ID - so content that changes every frame needs a
new texture per frame. Measured in process through the C ABI that is **0.31 ms**
end to end against 0.19 ms to apply an existing one. An earlier ~0.9-1.0 ms
figure was a measurement artefact of the PowerShell harness and is corrected in
[docs/cost_profile.md](docs/cost_profile.md). BlipBridge is best for a fixed set
of images reused many times, and workable for a video-like workload.

**`BB_ApplyTextureBatch` is not faster.** Measured at 10/50/100/200 Shapes it
lands within noise of the same number of individual calls, because each Shape
costs ~190 microseconds of real work and an ABI entry costs well under one. Use
it because one call is tidier, not because it is quicker. See
[docs/c_abi.md](docs/c_abi.md).

## Supported builds

| | |
|---|---|
| Platform | Windows x64 only |
| Host | PowerPoint (the accelerated path is refused elsewhere) |
| Office | **16.0.14334.20848** (PowerPoint LTSC 2021 x64), the build every offset was validated against |
| VBA | VBA7, 64-bit |
| macOS | **not supported** - see [docs/macos.md](docs/macos.md) |

On any other build, `BB_Init` returns `BB_E_UNSUPPORTED_BUILD` and
`BB_GetCapabilities` reports nothing. That is deliberate: the backend depends on
internal Office layouts, and guessing at an unvalidated one could corrupt a
document. This is **not** universal Office compatibility, and it is not marketed
as such.

Adding a build means re-validating its layouts, not editing a version number.

## Which Shapes work

Measured by creating a real instance of each class and applying to it, not by
reasoning about which ones expose a `Fill`:

| | Classes |
|---|---|
| **Native** | AutoShape, Freeform, TextBox, Placeholder, Callout, Group, group children, Picture, WordArt, Media |
| **Falls back to `UserPicture`** | Table, SmartArt, embedded OLE |
| **Refused** | Chart, Connector, Line |

Connectors and lines are refused *by name*, and that is the most important row.
They pass every **structural** check and fail the **semantic** one - two
independent layers, described in [docs/safety_model.md](docs/safety_model.md).
A Connector reports `Shape.Type = 1` (msoAutoShape) and presents byte-identical
internals to a rectangle - same wrapper, same FillFormat, same receiver - so no
structural check can tell them apart. Office's own `Fill.UserPicture` refuses one
with "value out of range", and a native apply **terminates PowerPoint**. The
guard is `Shape.Connector`, which is Office's own answer.

The full matrix, the evidence, and how a class earns native support are in
[docs/shape_compatibility.md](docs/shape_compatibility.md). It is a property of
one Office build and has to be re-run on any other.

## Safety

The backend calls undocumented Office internals, so every call is gated:

* **Module versions** - `oart.dll`, `ppcore.dll` and `gfx.dll` must all be the
  validated build.
* **Structural wrapper validation** - `Shape.Fill`'s PPCORE vtable is verified by
  *shape*, not address: its delegating thunks are decoded and the inner-object
  offset is read out of them rather than assumed.
* **Vtable identity** - every OART and GFX object is checked against its recorded
  vtable before a field is read.
* **Signature bytes** - the first sixteen bytes at every private entry point must
  match the bytes its ABI was derived from, or nothing is called.
* **No cached receiver** - the per-Shape receiver is re-resolved on every apply,
  because a deleted Shape still passes every pointer check.
* **Fail closed** - anything unrecognised is refused with a specific message, not
  worked around.

## Getting started

```text
MyGame.pptm
BlipBridge.dll        <- next to the presentation
```

1. Copy `BlipBridge.dll` beside your `.pptm`.
2. Import `vba/BlipBridge.bas` into the VBA project.
3. Use it:

```vb
Sub Demo()
    Dim shp As Shape

    If Not BlipBridge.IsAvailable Then
        MsgBox "BlipBridge: " & BlipBridge.Version   ' says why
        Exit Sub
    End If

    For Each shp In ActivePresentation.Slides(1).Shapes
        BlipBridge.UserPicture2 shp, "C:\textures\brick.png"
    Next shp
End Sub
```

The file is read and decoded once no matter how many Shapes get it, and a Shape
that already carries that image is skipped entirely. If you change a Shape's fill
by other means, call `BlipBridge.InvalidateShape shp` so the next call does real
work - see [docs/picture_cache.md](docs/picture_cache.md).

More in [examples/](examples/).

## Limitations

* Windows x64 and one Office build.
* Ten Shape classes are natively supported; Table, SmartArt and OLE fall back to
  `Fill.UserPicture`; Chart, Connector and Line are refused. See
  [the matrix](#which-shapes-work).
* Single-threaded: call from the thread that called `Initialize`.
* A texture handle must be released before PowerPoint exits. `Shutdown` and the
  add-in's teardown both do this; do not leak handles across a session.
* The batch API is a convenience, not a speed-up.
* Undo works, but each apply lands in the undo history like any other edit - a
  renderer doing thousands of applies will fill it.

## How it was built

Every claim in this README is backed by a measurement in [docs/](docs/):

| Document | What it establishes |
|---|---|
| [receiver_lookup.md](docs/receiver_lookup.md) | how a Shape reaches its internal fill receiver |
| [record_construction.md](docs/record_construction.md) | what the fill property record actually needs |
| [oart_abi.md](docs/oart_abi.md) | every private entry point, its ABI and its evidence |
| [native_texture.md](docs/native_texture.md) | reuse, lifetime, Undo, reference ownership |
| [safety_model.md](docs/safety_model.md) | structural against semantic validation, and why both are needed |
| [shape_compatibility.md](docs/shape_compatibility.md) | which Shape classes work, and the connector that crashes |
| [capabilities.md](docs/capabilities.md) | what each capability flag claims and why |
| [picture_cache.md](docs/picture_cache.md) | UserPicture2, the dispatch decision and both caches |
| [c_abi.md](docs/c_abi.md) | the public interface, handle and lifecycle contracts |
| [pixel_textures.md](docs/pixel_textures.md) | raw pixels, why mutation is unavailable, slowdown findings |
| [benchmarks.md](docs/benchmarks.md) | the numbers and how they were taken |
| [cost_profile.md](docs/cost_profile.md) | where each microsecond of an apply goes |
| [research.md](docs/research.md) | the journal, including the wrong turns |

Raw transcripts are in `docs/evidence/`.

## Building

Windows x64, MinGW-w64 UCRT and CMake:

```powershell
.\build.ps1 -Configuration Release
& ctest --test-dir build/Release --output-on-failure
```

`tests/abi_contract.cpp` runs without PowerPoint. The PowerPoint regressions in
`tools/` need a live host.

## License

MIT. See [LICENSE](LICENSE).

BlipBridge is not affiliated with or endorsed by Microsoft. It calls
undocumented internals of Microsoft Office, which may change or break in any
update; the version guards exist so that a change stops it rather than corrupts
anything.
