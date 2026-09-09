# The C ABI

`include/blipbridge/blipbridge.h` is the stable public interface. It is plain C:
no C++ types, no STL types, no exceptions and no allocation ownership cross it.
The COM Automation surface still exists, but it is now the research and
compatibility layer - nothing in normal use needs `regsvr32`, a ProgID,
`CreateObject`, or a COM add-in.

## Shape

```c
BB_Init / BB_Shutdown
BB_ApplyPicture                                       <- what most callers want
BB_InvalidateShape / BB_ClearPictureCache / BB_GetPictureCacheStats
BB_LoadTexture / BB_LoadTexturePixels
BB_ApplyTexture / BB_ApplyTextureBatch
BB_ReleaseTexture / BB_ClearTextures / BB_GetTextureCount
BB_GetCapabilities / BB_GetLastError
BB_GetVersion / BB_GetVersionString / BB_GetAbiVersion
```

Every entry point returns `BB_Result` (`0` success, negative failure). Handles
are `uint64_t`, opaque, and never zero when valid. On Windows x64 there is one
calling convention, so exports are undecorated - which is what lets VBA bind by
plain name:

```text
BB_ApplyTexture   BB_ApplyTextureBatch  BB_ClearTextures  BB_GetAbiVersion
BB_GetCapabilities BB_GetLastError      BB_GetTextureCount BB_GetVersion
BB_GetVersionString BB_Init             BB_LoadTexture    BB_LoadTexturePixels
BB_ReleaseTexture BB_Shutdown
```

## Versioning

`BB_ABI_VERSION` in the header is the compiled-against version; `BB_GetAbiVersion()`
is what the DLL implements. The VBA wrapper compares them in `Initialize` and
refuses to continue on a mismatch, so an old `.bas` paired with a newer DLL fails
with a clear message rather than calling something whose shape it has wrong.

It changes only when the exported surface stops being compatible - a changed
signature, a removed entry point, a changed meaning. **Adding** an export does
not bump it, because an older caller simply never calls the new one. That is why
`BB_LoadTexturePixels` and `BB_GetAbiVersion` arrived at ABI version 1.

**Version 2** was a changed meaning, not a new export. `BB_ApplyTexture` on a
Shape whose class has no native path used to return `BB_E_INVALID_ARG`; it now
returns `BB_E_UNSUPPORTED_SHAPE`, which says something different and useful - the
object is a perfectly good Shape, its class simply has no native picture-fill
path. A caller branching on `BB_E_INVALID_ARG` would silently stop matching, so
the version moved and the VBA wrapper refuses a mismatched pair loudly.

`BB_GetVersion` is separate: it is the release number and moves independently.

## The handle contract

`BB_Handle` is an opaque **64-bit** token on every platform. It is not a pointer.
Do not dereference it, cast it to one, or give it meaning through arithmetic. The
only guarantees are:

* zero is never a valid handle;
* values are never recycled within a process.

Lifecycle:

```text
BB_LoadTexture  or  BB_LoadTexturePixels
    -> BB_ApplyTexture, zero or more times, on any number of Shapes
    -> BB_ReleaseTexture
```

or `BB_ClearTextures`, or `BB_Shutdown`, either of which releases everything.

**Stale handles.** After release, a handle stays stale for the life of the
process. Using one returns `BB_E_INVALID_HANDLE`; because handles never recycle
it can never silently resolve to a different texture. Releasing twice returns
`BB_E_INVALID_HANDLE` on the second call rather than double-freeing.

In VBA a handle is a `LongLong`, matching the 64-bit contract exactly. It is
deliberately not `LongPtr`, which would imply pointer semantics the handle does
not have.

## Init and Shutdown

| Sequence | Behaviour |
|---|---|
| `BB_Init` twice | idempotent; re-probes and returns the same result |
| `BB_Shutdown` twice | both return `BB_OK` |
| `BB_Shutdown` before `BB_Init` | `BB_OK`, nothing to do |
| `BB_Shutdown` with live textures | releases them all first |
| `BB_Init` after `BB_Shutdown` | works; the library is usable again |
| any texture call before `BB_Init` | `BB_E_NOT_INITIALIZED` |

`BB_Init` claims the calling thread **even when the backend is unavailable**, so
a later call reports the real reason - `BB_E_UNSUPPORTED_BUILD`, say - rather
than the misleading `BB_E_NOT_INITIALIZED`.

`BB_Shutdown` releases only objects BlipBridge created. It never touches a Shape,
a receiver or any document object, because none of those is ever retained, so it
cannot reach through something Office has already destroyed. These sequences are
asserted in `tests/abi_contract.cpp`.

## Errors

`BB_GetLastError` is **copy-out**, not a pointer into internal storage:

```c
uint32_t BB_GetLastError(char* buffer, uint32_t capacity);
```

It writes UTF-8 into the caller's buffer and returns the size needed including
the terminator, so a caller can size a buffer by passing `(NULL, 0)` first. The
result is always terminated, including when truncated. There is no lifetime to
reason about, which is the point: an FFI caller should never have to know how
long a returned pointer stays valid.

The message describes the most recent failure **on the calling thread**.
`BB_GetVersionString` uses the same convention.

## Supported inputs

| | |
|---|---|
| Encoded images | whatever Office decodes - PNG and JPEG are tested |
| Raw pixels | BGRA32 only, stride >= width*4; see `pixel_textures.md` |
| Shape types | ten classes natively; others fall back or are refused - see shape_compatibility.md |
| Threading | the thread that called `BB_Init` |
| Office | the validated build only; others return `BB_E_UNSUPPORTED_BUILD` |

## Ownership

### Invalid against unsupported

Two refusals that used to look alike:

| Code | Meaning |
|---|---|
| `BB_E_INVALID_SHAPE` / `BB_E_INVALID_ARG` | the pointer is not a usable Shape - null, freed, not an `IDispatch`, or an object that will not answer `Shape.Type` |
| `BB_E_UNSUPPORTED_SHAPE` | a perfectly good Shape whose **class** has no native picture-fill path on this build |

The distinction matters because the two call for different responses. The first
is a bug in the caller or a Shape that died underneath it. The second is a fact
about PowerPoint, and the answer is `BB_ApplyPicture`, which routes such classes
to Office's own `Fill.UserPicture` where one exists.

## Ownership

| Thing | Rule |
|---|---|
| Texture handle | owned by the library until released; never recycled |
| Image bytes | copied before `BB_LoadTexture` returns |
| Shape pointer | **borrowed for the call only**, never stored |
| Error / version strings | copied into the caller's buffer |

**`BB_ApplyTexture` does not retain the Shape pointer.** Neither does the batch
form. Nothing in this library stores a Shape, a `FillFormat`, or a receiver
beyond the call that received it.

Passing `ObjPtr(shape)` is safe because the Shape is a live argument in the VBA
frame for the duration of the call, which keeps a reference alive. Never cache
the value of `ObjPtr` yourself.

The native side validates rather than trusts, in this order:

1. the pointer addresses committed memory;
2. it answers `QueryInterface` for `IDispatch`;
3. the object reports an AutoShape or Freeform `Type`;
4. its `Fill` is a PPCORE delegating wrapper, verified structurally;
5. the inner object presents the recorded OART FillFormat vtable;
6. the control block and receiver present theirs.

Steps 1 and 2 are cheap sanity checks and are **not** proof of the concrete type -
any COM object passes them. The structural Office validation in steps 4 to 6 is
the actual authority, and it is what distinguishes a `Shape.Fill` from a
`Shape.Line` or an unrelated object.

**Receiver pointers are never cached.** The receiver is re-resolved immediately
before every apply, because a deleted Shape still passes every pointer and vtable
check in the chain - see the hazard in `receiver_lookup.md`.

## Threading

Single-threaded apartment only. `BB_Init` records its thread; every later call is
checked and anything else returns `BB_E_WRONG_THREAD`. This is not caution: the
Office objects underneath use non-atomic reference counts.

## Failing closed

On a host where the accelerated backend cannot run, the library still loads and
still answers `BB_GetVersion`, `BB_GetCapabilities` and `BB_GetLastError`. The
texture entry points return `BB_E_UNSUPPORTED_HOST` or `BB_E_UNSUPPORTED_BUILD`
with a message that explains which. `tests/abi_contract.cpp` runs outside
PowerPoint precisely to hold that behaviour in place.

## Architecture

```text
    include/blipbridge/blipbridge.h   stable C ABI, platform independent
         |
    src/abi/blipbridge_abi.cpp        thread affinity, error text, translation
         |
    src/backend/backend.hpp           the platform seam
         |
   +-----+-----------------------------+
   |                                   |
windows_office_backend.cpp      unsupported_backend.cpp
   |                              (macOS and everything else)
src/backend/windows_office/
   oart_layout       module and layout guards, the validation cache
   native_apply      the property record, the transaction, the receiver call
   native_texture    the texture store: handles, ownership, lifetime
```

Nothing above `backend.hpp` knows Office exists. The reverse-engineered code is
reached only through `bb::Backend`, and its exceptions are converted to
`BackendResult` there so none can reach the C ABI.

`experiments/` holds only research now: the probes that produced the layouts, the
benchmarks, and the stage profiler. The dependency runs one way - a probe
includes `src/backend/windows_office/...`, and nothing under `src/` includes a
probe - so the shipping path cannot acquire instrumentation by accident.

## The batch API, measured

`BB_ApplyTextureBatch` takes parallel arrays - Shape pointers and handles - so no
C struct layout crosses the boundary and VBA can build both arrays trivially.

It works, and it is **not** a performance feature. Measured in process through
the real ABI, 50 iterations per size:

| Shapes | individual mean | batch mean | speed-up |
|---:|---:|---:|---:|
| 10 | 2.0081 ms | 2.0517 ms | 0.98x |
| 50 | 10.3810 ms | 9.8602 ms | 1.05x |
| 100 | 18.4012 ms | 18.2910 ms | 1.01x |
| 200 | 101.3182 ms | 107.2443 ms | 0.94x |

That is noise around 1.0. The reason is arithmetic rather than mysterious: each
Shape costs roughly 190 microseconds of real work - fetching `Fill`, re-resolving
the receiver, building the property record, committing the transaction - while an
ABI entry costs well under a microsecond. There is nothing meaningful to save.

A VBA caller saves a little more than this test can see, because each `Declare`
call also costs an interpreter-to-native transition that is absent here. But at
~190 microseconds of work per Shape, even a 5 microsecond transition is under 3%,
so **do not expect a batch speed-up**. Use it because one call is tidier than
two hundred, not because it is faster.

The 200-Shape row is slower per Shape for *both* legs (0.51 ms against 0.18-0.21
ms at smaller sizes). That is Office doing more work with 200 filled Shapes on a
slide, not a batch effect.
