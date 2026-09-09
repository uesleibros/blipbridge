# Changelog

All notable changes to BlipBridge. Dates are the day the work was validated on
the test machine.

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
