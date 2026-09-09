# macOS status

**There is no macOS backend, and this document exists so that nobody mistakes
that for an oversight.**

On macOS the library builds, loads, reports its version, and answers
`BB_GetCapabilities` with zero. Every texture entry point returns
`BB_E_UNSUPPORTED_HOST` with a message pointing here. It does not pretend, and it
does not ship a stub `.dylib` that quietly does nothing.

## Why the Windows backend does not port

The Windows implementation is not written against an Office API. It is written
against the *internals* of one Office build, and every part of that is
Windows-specific:

| Dependency | Why it does not transfer |
|---|---|
| `PPCORE.DLL`, `OART.DLL`, `GFX.DLL` | Windows PE modules; the Mac build has different binaries and a different internal structure |
| PPCORE delegating-wrapper vtables | a Windows C++ ABI vtable layout |
| OART/GFX vtable identity checks | validated against Windows RVAs |
| Thirteen private OART entry points | located by module RVA and verified by x86-64 instruction bytes |
| `GEL::ICachedImage::Create` | resolved from a Windows export table by MSVC-mangled name |
| The calling convention | Microsoft x64, including the hidden sret pointer convention |
| `IStream` | a COM interface the Mac pipeline does not use the same way |

None of that survives the move to a different binary format, a different
compiler ABI, and an ARM64 or x86-64 macOS build of Office. A port is new
research, not a recompile.

## What would have to be established first

Roughly the same ground the Windows work covered, in the same order:

1. **Is there an equivalent pipeline at all?** Does PowerPoint for Mac decode
   images through a cached-image abstraction that several Shapes can share, or
   does it decode per fill? If images are not shared internally, a texture handle
   has nothing to hold and the whole design is moot.
2. **Is there a reachable seam?** On Windows the entire approach rests on
   `Shape.Fill` being a thin wrapper delegating to an internal object. Whether an
   equivalent handle is reachable from the Mac automation surface is unknown.
3. **A stream-based image creator.** The Windows path works because GFX *exports*
   a creator taking a standard stream. Without an equivalent, bytes cannot reach
   the decoder without a temporary file, and `MemoryImageToFill` could not be
   claimed.
4. **A guard model that fails closed.** Whatever replaces module versions, vtable
   checks and signature bytes must be at least as strict. Applying an unvalidated
   layout to a user's document is worse than having no backend.
5. **Undo, ownership and persistence**, proven the same way they were on Windows:
   with a working control, attributed reference counts, and save/reopen.

## What a Mac contributor should know

The architecture is already prepared. `src/backend/backend.hpp` is the seam;
implementing `bb::Backend` and returning it from `CreateBackend()` is the entire
integration surface. Nothing above that line knows Office exists, and the C ABI,
the VBA wrapper and the tests need no changes.

`src/backend/unsupported_backend.cpp` is the current macOS answer and shows the
minimum an implementation has to provide.

The one thing not to do is weaken the guards to make something work. On Windows
every capability the library advertises is backed by a measurement in `docs/`,
and a Mac backend that advertised less rigorously would make the whole capability
system untrustworthy.
