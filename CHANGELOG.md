# Changelog

All notable changes to BlipBridge. Dates are the day the work was validated on
the test machine.

## [0.5.0] - 2026-09-10

Breaking: the public VBA API is now strongly typed and the ABI is version 3.
Callers written against 0.4.0 need the changes shown in the release notes.

### Fixed

- Allow release publication retries to replace matching assets on an existing
  release while preserving its notes and settings.
- Align release ZIP names with the workflow's `blipbridge-v<version>` convention
  and ship architecture-specific DLL names matching the VBA wrapper. Validate
  downloaded checksums before assembling the universal Windows package.
- Restore `windows.h` before `oleauto.h` in shared Automation and Office backend
  headers. Alphabetical include sorting broke MinGW's base type declarations,
  causing `EXTERN_C`, `DWORD`, and cascading `LCID` build errors. The formatter
  now prioritizes `windows.h` so subsequent formatting preserves this dependency.

## [Unreleased]

### Fixed

- **The VBA module did not compile.** `LoadImage`, `LoadImageFromFile`,
  `LoadTextureScaled` and `LoadTextureScaledFromFile` each declared the image
  request as `Optional ByRef request As BlipBridgeImageRequest`, and VBA does not
  allow a user-defined type as an optional parameter - so importing
  BlipBridge.bas failed with "invalid parameter type for optional parameter" and
  nothing in the module could be called. Shipped in 0.7.0 and found by opening
  the module, which is the one thing no gate did.

  The two image loaders are now `LoadImage` / `LoadImageFromFile` for a plain
  decode and `LoadImageProcessed` / `LoadImageProcessedFromFile` when a request
  is given. The two texture loaders take the request as a required parameter,
  since their names already say processing happens; `ImageRequest()` built with
  no arguments changes nothing.

- **`tools/check_vba_module.ps1`**, run by CI and by the release, checks the rules
  that made this possible: no user-defined type or array as an optional
  parameter, no optional without a default unless it is Variant, no required
  parameter after an optional one, no Object or Variant in a public signature,
  and every name the release gates grep for being a real declaration rather than
  a mention in a comment. Verified against the four broken declarations before
  being trusted - the first version of it silently passed two of them, because
  its parameter-list pattern stopped at the `()` of an array parameter.

## [0.7.0] - 2026-09-11

Image processing, and the resource split that makes it honest.

### Added - ABI version 5

- **A CPU image resource.** `BB_Image` in C, `BlipBridgeImage` in VBA: decoded
  BGRA that BlipBridge owns, so a picture can be cropped, oriented, scaled and
  warped repeatedly without decoding again.

  It exists *because* a texture is not one. `BlipBridgeTexture` stays exactly
  what it was - the Office resource, two opaque GFX pointers, no pixels anyone
  can reach - and does not start retaining a decoded copy behind your back, which
  would have doubled the memory of every texture in the process to serve the few
  that get processed twice. The two convert when you ask and never implicitly.

  `BB_LoadImage`, `BB_LoadImageFromFile`, `BB_LoadImagePixels`,
  `BB_GetImageSize`, `BB_ReleaseImage`, `BB_ClearImages`, `BB_GetImageCount`,
  `BB_CreateTextureFromImage`.

  In VBA that is `LoadImage` / `LoadImageFromFile` to decode and nothing else,
  and `LoadImageProcessed` / `LoadImageProcessedFromFile` to crop, orient or
  resize on the way in. Two names rather than one with an optional request,
  because VBA does not allow a user-defined type as an optional parameter.

- **Native decoding and processing.** Encoded images decode through Windows
  Imaging Component to canonical BGRA32 - straight alpha, because the resampler
  premultiplies where it needs to and doing it twice darkens transparent edges.
  `BB_ImageRequest` carries crop, transform, target size and filter through every
  entry point, with the stages in a fixed order: crop, then transform, then
  resize. Crop first because a region is named in the source's own coordinates;
  transform before resize so the target size always describes what comes out.

  Transforms are the lossless ones only - flip horizontal and vertical, rotate
  90/180/270 - where every output pixel is exactly one input pixel, so pixel art
  survives them. Arbitrary-angle rotation is absent on purpose: it needs
  resampling, a background colour and an output-size decision, none of which are
  free choices this library should make for you.

  Tested: PNG with and without alpha, JPEG, BMP. Windows decodes more and they
  will probably work, but an untested format is not a supported one.

- **`BB_LoadTextureEx` / `BB_LoadTextureFromFileEx`**: encoded straight to a
  texture with the same request, for a picture that needs processing once and no
  more. No CPU image is created, so nothing is retained beyond what Office holds.
  VBA: `LoadTextureScaled`, `LoadTextureScaledFromFile`.

- **`BB_WarpImageQuad` and `BB_ApplyImageQuad`**: a true projective mapping of an
  image onto four caller-supplied corners. The primitive returns a texture so a
  caller can decode once and warp as often as the quad moves; the convenience
  wrapper applies it in one call and owns the texture it makes.

  This is only possible because of what was measured first: **PowerPoint maps a
  picture fill linearly onto the Shape's bounding box and clips it to the path.**
  Seven quads, from rectangle to extreme trapezoid, agreed to within 0.8 source
  texels - the probe's own quantisation floor. Path geometry decides what is
  visible, not what is where. So warping into that bounding box makes Office's
  own mapping the identity, and the perspective survives.

  Verified twice: in isolation, where a rectangle is the exact identity, a
  parallelogram stays affine and an extreme trapezoid foreshortens monotonically;
  and in PowerPoint, where the source mid-row lands at 22% of a tapered quad's
  height rather than the 50% an affine map would give.

  Point order is documented and never reordered behind the caller. Nothing reads
  `Shape.Nodes` - the caller already knows the points, which is the premise.

  Bicubic is refused by name for the warp: sixteen taps per output pixel with a
  varying footprint is a cost that would be hidden rather than offered.

- **`BB_CAP_IMAGE_PIPELINE`** (0x0200) advertises the surface.

### Changed

- Public ABI 4 -> 5, moved in one step across the header, `BB_GetAbiVersion`,
  `BB_EXPECTED_ABI`, the wrapper, both CI gates, the release gate and the
  packaged-archive check - which now reads the header and wrapper back out of the
  zip and requires them to carry the new surface.

- Image handles come from their own numbering space, far from the texture store's,
  so a texture handle handed to an image call is refused rather than resolving to
  something unrelated. Neither space recycles a released handle.

### Research

- **How PowerPoint maps a picture fill**, which is what the quad pre-warp rests
  on: the picture is mapped approximately linearly across the Shape's bounding
  box and clipped to the freeform path. Seven quads from rectangle to extreme
  trapezoid agreed to a worst error of about **0.8 source texels**, which was the
  probe's own quantisation floor. Path geometry decides what is visible, not what
  is where - so warping into that bounding box first makes Office's final mapping
  the identity.

- **Dynamic textures are not shipped**, and this is what was established rather
  than assumed: `GEL::ICachedImage::Create` **copies** the pixel buffer it is
  given. Overwriting the source buffer after applying leaves the rendered Shape
  unchanged, and re-applying the same cached image - which distinguishes a stale
  repaint from a copy - leaves it unchanged too.

  That rules out the obvious externally-owned-buffer mutation path, by evidence.
  It does **not** prove that no mutable path exists anywhere inside GFX; deeper
  mutable internals remain future research. See `docs/image_pipeline.md`.

### Not validated

- **x86 Office runtime**, unchanged. The new image code is ordinary portable work
  with no Office in it, so it builds and its suites run on x86 - but that is a
  statement about image processing, and the x86 Office backend has still never
  run inside a real 32-bit PowerPoint. Building is not running.

## [0.6.0] - 2026-09-10

Adds the native ShapeRange apply and moves the public ABI to 4. Existing
`BB_ApplyTexture` and `BB_ApplyTextureBatch` calls are unchanged in behaviour;
callers written against ABI 3 need only the new header and wrapper.

### Fixed - correctness, independent of the new feature

- **A fill written by one API could be skipped by another.** `BB_ApplyTexture`
  changed a Shape's fill without telling `BB_ApplyPicture`'s cache, so a
  following `BB_ApplyPicture` for the image the Shape *used* to carry could skip
  as redundant and leave the wrong picture on screen. Every path that writes a
  fill now writes one shared per-Shape record - `BB_ApplyPicture`,
  `BB_ApplyTexture`, `BB_ApplyTextureIfChanged`, `BB_ApplyTextureRange` - and the
  ordinary apply only pays for the Shape key once something has asked for a skip.
  Nine assertions covering every order of those calls in
  `tools/test_range_apply.ps1`.

- **Freeforms could never be cached.** A Shape returned by
  `FreeformBuilder.ConvertToShape` reports a *ShapeRange* as its `Parent` rather
  than the Slide, and keeps doing so when fetched back out of the `Shapes`
  collection, so it had no readable slide id and no cache key - in either cache.
  Freeforms are a validated native class; they were uncacheable for a reason that
  has nothing to do with them. The key now tries one level up when the parent
  will not name a slide, and still refuses to key anything that names none.

- **`tools/test_native_texture_stress.ps1` guessed at live handles** by scanning
  the handle space for the first two that answer, which silently found nothing
  once a process had allocated past the end of its window. It loads its own.

### Added - ABI version 4

- **`BB_ApplyTextureRange(shapeRange, texture, applied)`**: one cached texture,
  one PowerPoint `ShapeRange`, one apply through Office's own range receiver.

  This is not a loop and not `BB_ApplyTextureBatch` with different arguments. A
  transaction carries no target - the receiver *is* the target - so a transaction
  holding several Shapes' fills cannot exist; what can exist is a receiver that
  stands for several Shapes, and `ShapeRange.Fill` is one. It presents the same
  PPCORE wrapper, the same OART FillFormat and the same validated receiver
  layout, so the existing structural walk accepts it unchanged.

  Measured through the public ABI in process, one image, against the same fills
  one Shape at a time: 32 Shapes in 1.9 ms against 6.6 ms, 100 in 5.5 ms against
  20.9 ms. `BB_ApplyTextureBatch` sits with the per-Shape leg, because it saves
  ABI crossings rather than Office work. Absolute figures move with machine load;
  the full table is in `docs/apply_fast_path.md`.

  - **All or nothing.** Every member is classified before any internal object is
    touched. One member without a validated native path - Connector, Line, Chart,
    Table - refuses the whole call, names which member, and never enters the
    private backend.
  - **One slide.** A PowerPoint `ShapeRange` cannot span slides, so neither can
    this. Several slides means one call per slide.
  - **Undo.** One range apply is one undo entry and one Undo reverts the whole
    fill, the same as Office's own `ShapeRange.Fill.UserPicture`. Filling the
    same Shapes one at a time leaves one entry each.
  - **Groups** may be members and are filled, and are never remembered for the
    skip, because a group's fill and its children's fills change each other.

  46 assertions in `tools/test_range_apply.ps1`. Reported through
  `BB_GetCapabilities` as `BB_CAP_RANGE_APPLY` (0x0100).

- **`BB_ApplyTextureIfChanged(shape, texture, skipped)`**: applies only when the
  Shape does not already carry that image. 0.022 ms against 0.189 - 8.5x - with
  no edit, no undo entry and no invalidation when it skips, and 4.9 us of
  overhead when it cannot. Images are compared by internal identity rather than
  by handle, and `Fill.Type` is re-read before any skip is granted. 33 assertions
  in `tools/test_apply_if_changed.ps1`.

- **`BlipBridge.ApplyTextureRange`** and **`BlipBridge.ApplyTextureIfChanged`**
  in the VBA wrapper, strongly typed against `PowerPoint.ShapeRange` and
  `PowerPoint.Shape`.

### Changed

- **Public ABI 3 -> 4**, because `BB_ApplyTextureRange` is a finalised public
  export. `BB_GetAbiVersion`, `BB_EXPECTED_ABI` in the VBA wrapper, the CI gate
  and the release gate all move together, and the release now refuses to publish
  an ABI 4 build that does not export the range API and define
  `BB_CAP_RANGE_APPLY` as 0x0100.

- The Office module handles are resolved once and **pinned** rather than looked
  up on every apply. Three `GetModuleHandleW` calls, each taking the loader lock,
  were 7.1 of the receiver resolution's 18.6 microseconds. A pinned module cannot
  be unloaded, so its base cannot move - a stronger guarantee than the re-lookup
  it replaces, which could only notice a swap after it had happened.

- Release validation now covers range operations: the heterogeneous-class range,
  groups in a range, a second texture over the same range, the same-slide
  restriction, undo and redo of a whole range, and every order of the three
  fill-writing APIs against the shared record.

### Not validated

- **x86 Office runtime.** The x86 DLL is built and gate-checked in CI - exports,
  architecture, calling convention, packaging - and the C ABI contract test runs
  on it. It has never been run inside a real 32-bit PowerPoint, and the x86
  backend refuses by design rather than implementing the native path. Building
  is not running; see `docs/windows_x86.md`.


### Fixed

- **`UserPicture2` broke after `ClearTextures`.** The texture store owned each
  image outright and handed callers a handle into its map; the picture cache took
  one of those handles, so clearing the map - or releasing that handle - destroyed
  the image the cache still pointed at, and the next call failed with "Texture
  handle 16777216 is not valid (it was released)". A public handle is now an
  external token that *references* an image, and the handle table and the picture
  cache own it independently. `ClearTextures` drops handles, `ClearPictureCache`
  drops the cache's images and Shape skip state, neither can reach the other's,
  and `BB_Shutdown` clears both. No retry, no handle recycling: released handles
  remain permanently stale. Covered by `tools/test_cache_ownership.ps1`.

### Added - ABI version 3

- **`BB_LoadTexturePixelsScaled`**: raw BGRA32 resampled to any size with an
  explicit filter - `BB_SCALE_NEAREST`, `BB_SCALE_BILINEAR`, `BB_SCALE_BICUBIC`.
  Every filter named is implemented and tested; there are no placeholders.
  Upscaling and downscaling at arbitrary ratios, padded source strides, correct
  premultiplied-alpha interpolation so transparent edges do not darken, a
  bit-exact fast path when the sizes match, and full 64-bit validation of every
  size and product before anything is read or allocated. Capability bit
  `BB_CAP_SCALED_PIXELS`. See `docs/resampling.md`.
- `tests/resample_contract.cpp` runs in CI on both architectures with no Office,
  including nearest's exact byte output, and `bb_resample_benchmark` times every
  filter.

### Changed - the VBA API is strongly typed

- Shapes are `PowerPoint.Shape` and handles are `BlipBridgeTexture` - a public
  two-Long type identical on both architectures. No `Object`, no `Variant` and no
  `LongLong` in any public signature; `LongLong` does not exist in 32-bit VBA and
  a `Double` would lose precision above 2^53, so two Longs are the one exact
  shape both can express. The native handle stays an opaque `uint64_t`, and the
  per-architecture marshalling is private.
- `BlipBridgeScaleFilter` is a real enum, with `LoadTexturePixelsNearest`,
  `...Bilinear` and `...Bicubic` convenience wrappers over the same native call.
- **ABI version 3.** The wrapper refuses a DLL that reports anything else, and CI
  now asserts the header and the wrapper agree.
- The build emits `BlipBridge-x64.dll` and `BlipBridge-x86.dll` directly, so a
  local build is the file the wrapper looks for rather than something the
  packaging step renames.

### Measured

- Resampling, milliseconds, **scaling only** - cached-image creation and the
  apply are measured separately: 512x512 → 1920x1080 costs 2.09 (nearest), 47.5
  (bilinear), 186 (bicubic). Nearest is roughly twenty times cheaper than
  bilinear and eighty times cheaper than bicubic. Hoisting the per-column weights
  out of the pixel loop took bicubic from 267 ms to 186 ms. No SIMD or threading
  was added; nothing has shown them necessary.
- That cost is paid on **every** `BB_LoadTexturePixelsScaled` call. It amortises
  to nothing if you keep the handle and reuse it, and is paid in full by any
  workload that builds a new scaled texture each time.

## [0.4.0] - 2026-09-10

First public release. `uesleibros/blipbridge`, MIT, with CI and an automatic
release pipeline.

### Added - distribution

- **Windows x86 as a build target, and deliberately not a supported backend.**
  CMake decides architecture from the toolchain and ships different file sets:
  x64 gets the validated PowerPoint backend, x86 gets the public C ABI over a
  backend that refuses. The reverse-engineered sources are not compiled on x86 at
  all, because they would compile and be wrong - they depend on the x64 calling
  convention, per-build module RVAs, x64 instruction bytes for the signature
  checks, and pointer-sized record slots. See `docs/windows_x86.md`.
- **One VBA wrapper for both architectures.** `LongLong` exists only in 64-bit
  Office, so handles cross as two `ByVal Long` on x86 - exactly how a stdcall
  frame carries a 64-bit value. The public API takes and returns `Variant`, so
  caller code is identical on both.
- **Wrong-package detection.** `ERROR_BAD_EXE_FORMAT` becomes "Architecture
  mismatch: ... is the 64-bit build, but this PowerPoint is 32-bit", naming the
  package to download, instead of error 193.
- **CI** across x64/x86 and Release/Debug: configure, build, confirm the DLL
  really is the architecture it claims, test, and check the export surface.
- **Release workflow** on a `v*` tag: build, test, package, SHA-256, publish.
  Nothing publishes if a build or test job fails.
- `SECURITY.md` and four issue templates. The crash and unsupported-build
  templates ask for version, build, architecture and capability output, and both
  say not to attach a memory dump - those carry whatever document was open.

### Changed

- **x86 exports are `__stdcall`**, with `-Wl,--kill-at` keeping the names
  undecorated. `BB_CALL` was empty, which on 32-bit means cdecl, and VBA's
  `Declare` can only call stdcall - a cdecl export would bind and then unbalance
  the stack on every call. x64 is unaffected.
- The library version lives in the public header alone. It had drifted to three
  different values across the header, the ABI implementation and CMake.

### Fixed

- The ABI contract test declared its function pointers without a calling
  convention, which is cdecl on x86 against a stdcall DLL. Found by CI as a
  segfault with no useful message.

## [Unreleased]

### Changed - ABI version 2

- `BB_ApplyTexture` on a Shape whose class has no native path now returns
  **`BB_E_UNSUPPORTED_SHAPE`** where it returned `BB_E_INVALID_ARG`. That is a
  changed meaning on an existing entry point, so `BB_ABI_VERSION` moved to 2 and
  the VBA wrapper refuses a mismatched pair loudly. `BB_E_INVALID_SHAPE` now means
  only what it says: the object is not a usable Shape.

### Added - the semantic safety layer

- **One authority for Shape eligibility**, `ClassifyShapeForNativePictureFill` in
  `src/backend/windows_office/shape_policy.cpp`, returning `NativeSupported`,
  `FallbackSupported`, `Unsupported` or `Invalid`. The C ABI, the COM surface and
  the `UserPicture2` dispatcher all ask it; none re-derives it. Documented in
  `docs/safety_model.md`: structural validation proves an object is the Office
  object we expect, semantic validation proves the operation is meaningful for
  that Shape class, and the Connector proves neither implies the other.
- **Permanent regressions** (`tools/test_semantic_guards.ps1`) for Connector,
  Line and WordArt, each in its own PowerPoint. They assert something stronger
  than "it did not crash": the backend counts entries into the private OART
  apply, and Connector and Line report `0 -> 0` - the dangerous call was never
  reached. WordArt reports `0 -> 1` and is verified end to end: still WordArt,
  text and geometry intact, picture fill persists through save, reopen, undo and
  redo.
- **Lifecycle cache regressions** (`tools/test_shape_lifecycle_cache.ps1`) across
  Duplicate, Copy/Paste, Group, Ungroup, move between slides, delete, undo
  delete, redo delete and reopen.
- The compatibility matrix now records structural and semantic verdicts as
  separate columns, plus route, `Fill.Type` before and after, undo, redo, reopen
  and host survival, and uses `CrashedDuringResearch` for a class whose host died.

### Fixed

- **The picture path could fall back for the wrong reasons.** It treated any
  refusal as "this class needs the fallback", which would have routed a deleted or
  unusable Shape - and a Connector - into `Fill.UserPicture` instead of reporting
  the real problem. It now switches on the explicit verdict, so only
  `FallbackSupported` falls back.
- **Groups and their children are no longer cached by the skip cache.** Filling a
  group changes what its children render - `tools/probe_group_fill_propagation.ps1`
  measures a child's PNG at 33,515 bytes before and 35,725 after, with `Fill.Type`
  reading 6 throughout - so a child's remembered texture went stale the moment its
  group was filled, and a later apply would have been skipped, leaving the wrong
  picture on screen silently. Caught by the new lifecycle suite.

### Measured

- The semantic guard costs two Automation property reads, about two microseconds.
  Apply median 0.1547 ms against 0.1673 ms before it existed - within machine
  variance, so no measurable hot-path cost.
- A group child *can* be keyed safely, contrary to the previous comment in the
  code: `tools/probe_shape_identity.ps1` shows its `Parent` is the Slide, it
  reports a `SlideID`, its `Id` collides with nothing on the slide, and Ids
  survive grouping and ungrouping. It is excluded from the cache for the
  propagation reason above, not for want of a key.

### Added

- **`BB_ApplyPicture` / `UserPicture2`**: one call taking a Shape and a file
  path, which chooses the accelerated path, Office's own `Fill.UserPicture`, or a
  specific refusal, so a caller never has to know the Shape's class. With
  `BB_InvalidateShape`, `BB_ClearPictureCache` and `BB_GetPictureCacheStats`, and
  the capability bit `BB_CAP_APPLY_PICTURE`. See `docs/picture_cache.md`.
- Two caches under it: **path to texture**, keyed on path plus size plus
  last-write time so an edited file is never served stale, and **Shape to last
  texture**, so repeating an image on a Shape does no Office work. Measured at
  0.49 ms against 6.29 ms for a real apply through the same boundary - a gap
  wider than the raw apply cost, because a real fill change also makes PowerPoint
  repaint.
- Specific error codes where a generic one used to do: `BB_E_UNSUPPORTED_SHAPE`,
  `BB_E_FILE_NOT_FOUND`, `BB_E_FALLBACK_FAILED`.
- **Ten Shape classes are natively supported**, up from two: AutoShape, Freeform,
  TextBox, Placeholder, Callout, Group, group children, Picture, WordArt and
  Media. Table, SmartArt and OLE fall back; Chart is unfillable either way.
  `docs/shape_compatibility.md` has the matrix and what each class had to survive.

### Fixed

- **A native apply to a Connector terminated PowerPoint.** A Connector reports
  `Shape.Type = 1` (msoAutoShape) and presents byte-identical internals to a
  rectangle, so neither the type check nor the structural walk separated them.
  Office's own `Fill.UserPicture` refuses one, from a check PPCORE performs before
  the fill handler is reached; the native apply reproduces the handler and not
  that pre-check. Connectors and lines are now refused by `Shape.Connector`, and a
  Shape that will not answer the question is refused rather than assumed safe.

### Measured

- **Stage-by-stage cost profile** (`docs/cost_profile.md`,
  `tools/run_stage_profiler.ps1`). About 80% of an apply is inside Office's own
  receiver call - the document edit, the undo entry, the invalidation. Record
  construction, the image sub-record, both AddRef steps, the stretch holder, the
  transaction constructor and all four destructors together are **under 4%**,
  about 8 microseconds, so reusing or pooling any of those allocations cannot
  pay. That is now measured rather than assumed.
- **The ~0.9-1.0 ms new-content figure was a harness artefact** and is corrected
  everywhere it appeared. It was taken from PowerShell, which pays a
  cross-process Automation round trip per call and made three per frame. In
  process through the public C ABI a new-content frame costs **0.31-0.34 ms**.
- **A pool of pre-created images is safe but is not a speed-up.** Eight distinct
  images applied in turn over 400 applies: reference counts bounded, private
  bytes +2-3 MB, apply cost identical to reusing one texture. Without mutation a
  pool recycles handles, not work, so no pooling API was added.

### Changed

- **Guard chain roughly halved**, from 0.042 ms to 0.022 ms per apply, with every
  check kept exactly as strict. `SynchroniseOfficeModules()` fetches the three
  module handles once per operation instead of six times, and `SizeOfImage` is
  cached per module handle alongside the version check that was already cached.
  Nothing was removed or relaxed: the receiver is still re-resolved on every
  apply, every vtable checked, every signature verified. Hot-path apply median
  0.176 ms -> 0.167 ms.
- **The validated Windows backend moved into the production tree.**
  `src/backend/windows_office/` now holds `oart_layout`, `native_apply`,
  `native_texture` and a `native_texture.hpp` seam; `experiments/` keeps only
  research - the probes, the benchmarks and the stage profiler. The production
  backend no longer includes the research header, and the two research entry
  points that lived in `native_apply.cpp` moved to
  `experiments/exp_internal_blip/native_apply_probe.cpp`. Behaviour is unchanged:
  the reference-count timelines the Office harnesses print are identical before
  and after the move.
- Removed four empty placeholder directories (`src/core`, `src/diagnostics`,
  `src/office`, `experiments/exp_stream`) that described a hypothetical
  architecture rather than the real one.

## [0.3.0] - 2026-09-09

### Added

- `BB_LoadTexturePixels`: build a texture from a raw BGRA32 buffer, skipping
  image decoding. Capability bit `BB_CAP_RAW_PIXELS`, and
  `BlipBridge.LoadTexturePixels` in the VBA wrapper.
- `BB_GetAbiVersion` and `BB_ABI_VERSION`, with the VBA wrapper refusing to run
  against a DLL whose ABI version it does not match.
- CI (`.github/workflows/windows.yml`): Release and Debug builds, CTest, and an
  export-surface check so a renamed export cannot silently break the wrapper.
- `docs/pixel_textures.md`.

### Measured, and deliberately not marketed

- **Raw pixels are ~15-18% slower** than loading an encoded PNG of the same
  dimensions (0.0280 ms against 0.0240 ms at 64x64; 0.2837 against 0.2315 at
  256x256). The path exists so a caller holding pixels need not encode a PNG
  first, not because it is faster.
- **Image mutation is not available.** No GFX export writes to an existing
  `IImage` or `ICachedImage`, and `ICachedImage` carries a unique ID and a
  caching flag - the marks of a deliberately immutable, content-addressed object.
  Changing content therefore means a new texture per frame, ~0.9-1.0 ms end to
  end against 0.19 ms to apply an existing one.
- **The surface-format argument carries no information.** Every value from 0 to
  24 produced an identical correct BGRA render, with the argument mapping proved
  separately via stride. Only BGRA32 is offered rather than a parameter that
  looks like a contract and is not.
- **No slowdown over time reproduced**, across three configurations totalling
  thousands of applies. Calling `StartNewUndoEntry` per apply - the obvious
  looking fix - makes things slower, not faster.

### Documented

- Handle contract: opaque 64-bit token, never a pointer, never recycled, exact
  stale-handle behaviour.
- Init/Shutdown semantics for every ordering, asserted in the ABI tests.
- `BB_GetLastError` is copy-out with no lifetime to reason about.
- `BB_ApplyTexture` does not retain Shape pointers; receivers are never cached.

## [0.2.0] - 2026-09-09

The release that turns a validated backend into a usable library.

### Added

- **Public C ABI** (`include/blipbridge/blipbridge.h`): `BB_Init`,
  `BB_Shutdown`, `BB_LoadTexture`, `BB_ApplyTexture`, `BB_ApplyTextureBatch`,
  `BB_ReleaseTexture`, `BB_ClearTextures`, `BB_GetTextureCount`,
  `BB_GetCapabilities`, `BB_GetLastError`, `BB_GetVersion`,
  `BB_GetVersionString`. Plain C, fixed-width types, opaque handles, explicit
  error codes, no exceptions crossing the boundary.
- **VBA wrapper** (`vba/BlipBridge.bas`) that loads the DLL from beside the
  presentation with `LoadLibraryW`, so deployment needs no registration, no
  ProgID and no add-in install.
- **Platform seam** (`src/backend/backend.hpp`) with a Windows implementation and
  an honest unsupported-platform implementation, so macOS gets a documented
  refusal rather than a fake.
- **Batch apply**, and the measurement showing it is a convenience rather than a
  speed-up.
- **ABI contract tests** (`tests/abi_contract.cpp`) that run without PowerPoint.
- Docs: `c_abi.md`, `capabilities.md`, `macos.md`.

### Changed

- Capabilities are computed at runtime from the live modules instead of being
  hard-coded, so they describe the machine the library is actually on.
- `LoadTexture` decodes into a reusable native texture instead of returning
  `E_NOTIMPL`. `SetImageBytes` still returns `E_NOTIMPL`.
- COM Automation is now the research and compatibility layer rather than the
  primary interface. It still works and its tests still run.

### Notes

- `FillOnly` stays `False` deliberately; see `docs/capabilities.md`.
- `BB_ApplyTextureBatch` measured within noise of individual calls at
  10/50/100/200 Shapes.

## [0.1.0] - 2026-09-09

The research that made a native backend possible, and then made it work.

### Added

- Native cached-image apply: bytes to an in-memory stream, to PowerPoint's own
  exported cached-image creator, to its own property record, transaction and
  receiver. No `Fill.UserPicture`, no donor Shape, no `PickUp`/`Apply`, no file.
- Reusable texture handles with attributed lifetime, proven to release fully on
  document close.
- Guard model: module version checks, structural PPCORE wrapper validation,
  OART/GFX vtable checks, per-function signature bytes, receiver re-resolution on
  every apply, fail-closed on unknown layouts.
- Donor `PickUp`/`Apply` fallback, still present and still tested.

### Measured

- `ApplyTexture` 0.1860 ms mean against `Fill.UserPicture` 0.6588 ms - about
  3.5x, with a wider margin at p99 (0.3361 ms against 1.6002 ms).
- One decode serves any number of Shapes: 10,000 applies added zero
  cached-image creations.
- 12 `UserPicture` calls write 24 `Content.MSO` files; 12 native applies write
  none.
- Undo and Redo both work, verified against a control that undoes an ordinary
  `Fill.UserPicture` first.
- Reference counts return to exactly 1 - the handle's own - once documents close.

### Corrected along the way

Each of these was believed and then disproven by measurement; the journal in
`docs/research.md` keeps the wrong turns visible.

- "The receiver is per Shape" was under-determined until one Shape was filled
  twice in the same experiment.
- `receiver+0x20` is a global allocation counter, not a Shape index.
- Native applies were expected not to be undoable. They are; the earlier test
  used a presentation with no window, so its control failed too.
- The first benchmark showed the native path 2.5x *slower*, which turned out to
  be per-apply validation rather than the work, and is now cached by module.
