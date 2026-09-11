<p align="center">
  <img src="assets/blipbridge-logo.png" alt="BlipBridge logo" width="240">
</p>

<h1 align="center">BlipBridge</h1>

<p align="center">
  Fast, reusable image textures for ordinary PowerPoint Shapes, callable from VBA.
</p>

<p align="center">
  <a href="https://github.com/uesleibros/blipbridge/actions/workflows/ci.yml"><img src="https://github.com/uesleibros/blipbridge/actions/workflows/ci.yml/badge.svg" alt="CI build status"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/license-MIT-22c55e" alt="License: MIT"></a>
  <a href="#building"><img src="https://img.shields.io/badge/C%2B%2B-20-2563eb" alt="C++20"></a>
  <a href="#supported-builds"><img src="https://img.shields.io/badge/native_backend-Windows_x64-7c3aed" alt="Native backend: Windows x64"></a>
</p>

```vb
BlipBridge.UserPicture2 shp, "C:\textures\brick.png"   ' that is the whole API
```

It decides the rest: the accelerated path for Shape classes that have one,
Office's own `Fill.UserPicture` for classes that do not, the file read and
decoded once, and no work at all when that Shape already carries that image. For
full control the texture handles are still there:

```vb
Dim texture As BlipBridgeTexture
Dim shp As PowerPoint.Shape

texture = BlipBridge.LoadTexture(bytes)   ' decode once
BlipBridge.ApplyTexture shp, texture      ' ~0.19 ms, no file, no donor Shape
BlipBridge.ReleaseTexture texture
```

Filling many Shapes with the *same* texture is one call, and one Office edit
rather than N:

```vb
Dim names As Variant
Dim rng As PowerPoint.ShapeRange
Dim applied As Long

names = Array("face1", "face2", "face3", "face4")
Set rng = Slide1.Shapes.Range(names)

applied = BlipBridge.ApplyTextureRange(rng, grassTexture)
```

Every Shape in the range gets that one cached texture, the whole thing counts as
**one** PowerPoint edit - one Ctrl+Z takes it back - and the range must belong to
one slide, because a PowerPoint `ShapeRange` cannot span slides. One member with
no validated native path, a Connector or a Chart say, refuses the whole call by
name before anything internal is touched.

Reach for something else when:

- **each Shape needs a different texture** - `ApplyTextureBatch`, which takes one
  handle per Shape. It saves the call overhead, not Office's work.
- **there is one Shape** - `ApplyTexture`. At one Shape the range path costs the
  same and says less.

Skip the work entirely when nothing would change:

```vb
If BlipBridge.ApplyTextureIfChanged(shp, texture) Then
    ' the Shape already had this image; Office was not touched at all
End If
```

The public API is strongly typed on both architectures: `PowerPoint.Shape` for
Shapes and `BlipBridgeTexture` for handles, with no `Object`, no `Variant` and no
`LongLong` in any public signature. Raw BGRA pixels can also be resampled on the
way in, with an explicit filter:

```vb
texture = BlipBridge.LoadTexturePixelsScaled( _
    pixels, srcW, srcH, srcStride, dstW, dstH, BBScaleBicubic)
```

That chooses how BlipBridge resamples the image **before** Office receives it. It
does not change how Office draws it - see [docs/resampling.md](docs/resampling.md).

No `regsvr32`. No ProgID. No `CreateObject`. No add-in installer. Put the DLL
next to your presentation and import one `.bas` module.

## Status

Two different things, kept apart on purpose. **A green build says nothing about
whether the PowerPoint backend is safe** - hosted CI runners have no Office at
all.

| | Builds in CI | Native PowerPoint backend |
|---|---|---|
| **Windows x64** | yes | **validated** on Office 16.0.14334.20848 |
| **Windows x86** | yes | **not validated** - never run in a real 32-bit PowerPoint; loads and refuses every texture call |
| macOS | no | not implemented, and not a port |

The x86 package is real and useful for testing the ABI, the wrapper and the
packaging on 32-bit Office. It will not accelerate anything: the 64-bit backend
is not even compiled into it, because code that links and is wrong is the worst
possible outcome for a library that drives undocumented Office internals. See
[docs/windows_x86.md](docs/windows_x86.md).

Backend validation happens on a machine with the validated Office build, and the
transcripts are committed under [docs/evidence/](docs/evidence/) so the evidence
travels with the repository even though CI cannot reproduce it.

> **Research-grade, working, and narrow.** The accelerated backend runs against
> **one validated Office build** and refuses anything else. Read
> [Supported builds](#supported-builds) before depending on it.

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
tail. Measured on the validated test environment - Office 16.0.14334.20848,
Windows x64, one machine, one workload. Performance varies with Office build,
hardware, image size and how much else the document is doing. One-off costs: `LoadTexture` 0.0231 ms, `ReleaseTexture` 0.0001 ms.
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

**Filling a ShapeRange is one Office edit, not N.** `ApplyTextureRange` goes
through Office's own range receiver, so the per-Shape cost falls from ~0.19 ms
to well under a tenth of that at useful sizes:

| Shapes | one at a time | `ApplyTextureRange` | speed-up |
|---:|---:|---:|---:|
| 1 | 0.224 ms | 0.233 ms | 1.0x |
| 8 | 1.77 ms | 0.63 ms | 2.8x |
| 32 | 6.65 ms | 1.88 ms | 3.5x |
| 100 | 20.9 ms | 5.55 ms | 3.8x |

One range apply is one undo entry, the same as Office's own
`ShapeRange.Fill.UserPicture`. A ShapeRange cannot span slides, so several
slides means one call per slide. One member with no validated native path - a
Connector, a Line, a Chart - refuses the whole call before anything internal is
touched. See [docs/apply_fast_path.md](docs/apply_fast_path.md).

**`BB_ApplyTextureBatch` is not faster**, and is a different thing. It takes an
array of Shapes and still makes one Office edit per Shape, so it lands within
noise of the same number of individual calls: it saves ABI crossings, not work.
Use it because one call is tidier; use `ApplyTextureRange` when you want it
quicker. See [docs/c_abi.md](docs/c_abi.md).

**Repeating an apply that changes nothing is nearly free.**
`ApplyTextureIfChanged` costs 0.022 ms against 0.189 when the Shape already
carries the image - 8.5x - and 4.9 microseconds more than `ApplyTexture` when it
cannot skip.

## Supported builds

| | |
|---|---|
| Platform | Windows. x64 has a validated backend; x86 builds and refuses |
| Host | PowerPoint (the accelerated path is refused elsewhere) |
| Office | **16.0.14334.20848** (PowerPoint LTSC 2021 x64), the build every offset was validated against |
| VBA | VBA7. One `.bas` serves both 32-bit and 64-bit Office |
| macOS | **not supported** - see [docs/macos.md](docs/macos.md) |

The validated build is a property of **that build**, not of a version family.
`16.0.14334.x` is not assumed to work; only the exact build was tested, and
anything else fails closed. Adding one means re-validating its layouts against a
real install - use the
[unsupported-build issue template](https://github.com/uesleibros/blipbridge/issues/new?template=unsupported_office_build.yml).

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

## Textures, caches and ownership

Three things a caller should know, and they are the difference between the API
behaving obviously and behaving mysteriously.

**A texture handle is an opaque token.** Not a pointer, not an index, and no
arithmetic on it means anything. Zero is never valid. Handles are never recycled,
so a released one is permanently stale and can never silently resolve to a
different image.

**Two owners, independently.** A handle *references* an image; it does not own it
outright. `UserPicture2`'s cache holds its own reference to the images it created.
An image lives until the last owner lets go:

| Call | What it drops |
|---|---|
| `ReleaseTexture` | that one handle's reference |
| `ClearTextures` | every caller-visible handle |
| `ClearPictureCache` | the `UserPicture2` cache's own images, and its per-Shape skip state |
| `Shutdown` | both, because it is the only call that should |

`ClearTextures` cannot break `UserPicture2`, and `ClearPictureCache` cannot break
a handle you still hold. Neither touches resources the other owns.

**`InvalidateShape`** forgets what `UserPicture2` last put on one Shape. Call it
after changing that Shape's fill by any other means, so the next call does real
work instead of correctly-but-wrongly skipping.

Full detail, including the defect that produced this model, is in
[docs/picture_cache.md](docs/picture_cache.md).

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

### Which package?

Check **File > Account > About PowerPoint**. The first line ends with `64-bit` or
`32-bit`, and that is the package to download from
[Releases](https://github.com/uesleibros/blipbridge/releases):

```text
blipbridge-<version>-windows-x64.zip     64-bit PowerPoint
blipbridge-<version>-windows-x86.zip     32-bit PowerPoint
```

Get it wrong and `BlipBridge.bas` says so in words - "Architecture mismatch: ...
is the 64-bit build, but this PowerPoint is 32-bit" - rather than leaving you
with error 53, error 193, or "Bad DLL calling convention".

Verify a download with `sha256sum -c SHA256SUMS.txt`.

### Install

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
| [picture_cache.md](docs/picture_cache.md) | UserPicture2, the ownership model and both caches |
| [resampling.md](docs/resampling.md) | the scaling filters, their semantics and their cost |
| [c_abi.md](docs/c_abi.md) | the public interface, handle and lifecycle contracts |
| [pixel_textures.md](docs/pixel_textures.md) | raw pixels, why mutation is unavailable, slowdown findings |
| [benchmarks.md](docs/benchmarks.md) | the numbers and how they were taken |
| [cost_profile.md](docs/cost_profile.md) | where each microsecond of an apply goes |
| [windows_x86.md](docs/windows_x86.md) | what x86 is today, and what validating it would take |
| [research.md](docs/research.md) | the journal, including the wrong turns |

Raw transcripts are in `docs/evidence/`.

## Building

MinGW-w64 and CMake. The architecture comes from the toolchain, so a UCRT64 shell
produces x64 and a MINGW32 shell produces x86 - there is no flag to get wrong,
and CMake prints which one it chose.

```powershell
.\build.ps1 -Configuration Release
& ctest --test-dir build/Release --output-on-failure
```

or directly:

```bash
cmake -S . -B build/Release -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build build/Release --parallel
ctest --test-dir build/Release --output-on-failure
```

## Testing

Two suites, and the difference matters:

| Suite | Needs PowerPoint | Runs in CI |
|---|---|---|
| `ctest` - ABI and COM contracts | no | yes, both architectures |
| `tools/test_*.ps1` - the Office regressions | **yes** | **no** |

The Office suites cover the semantic Shape guards, the compatibility matrix, the
per-class stress runs, the picture cache, Shape lifecycle, undo and redo. They
need a live PowerPoint on the validated build, so they are never run or reported
by CI. Run them locally:

```powershell
.\tools\test_semantic_guards.ps1        # Connector, Line, WordArt
.\tools\test_shape_compatibility.ps1    # the matrix, one process per class
.\tools\test_shape_class_stress.ps1     # 1000 applies per native class
.\tools\test_picture_cache.ps1          # UserPicture2 and both caches
.\tools\test_shape_lifecycle_cache.ps1  # duplicate, group, delete, undo
```

See [.github/workflows/README.md](.github/workflows/README.md) for the full split
and how an Office-capable runner could be added later.

## Releasing

Tag with SemVer and the release workflow does the rest - build both
architectures in Release and Debug, test, package, checksum, and publish:

```bash
git tag v0.4.0 && git push origin v0.4.0
```

Nothing is published if a build or test job fails.

## Contributing

[CONTRIBUTING.md](CONTRIBUTING.md) covers the workflow. Two things are worth
knowing before opening a pull request:

* **Guards are not optional.** Version checks, structural validation, signature
  checks, semantic Shape eligibility and per-apply receiver re-resolution all
  exist because of specific measured failures. Weakening one needs a better
  reason than speed.
* **Evidence over reasoning.** A Shape class is not supported because it looks
  like it should be - a Connector looks identical to a rectangle all the way down
  and terminates PowerPoint. Support comes from running the harnesses.

Security issues, and anything that could damage a document, go through
[SECURITY.md](SECURITY.md) privately first.

## License

MIT. See [LICENSE](LICENSE).

BlipBridge is not affiliated with or endorsed by Microsoft. It calls
undocumented internals of Microsoft Office, which may change or break in any
update; the version guards exist so that a change stops it rather than corrupts
anything.
