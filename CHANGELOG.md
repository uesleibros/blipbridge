# Changelog

All notable changes to BlipBridge. Dates are the day the work was validated on
the test machine.

## [Unreleased]

### Fixed

- Restore `windows.h` before `oleauto.h` in shared Automation and Office backend
  headers. Alphabetical include sorting broke MinGW's base type declarations,
  causing `EXTERN_C`, `DWORD`, and cascading `LCID` build errors. The formatter
  now prioritizes `windows.h` so subsequent formatting preserves this dependency.

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
