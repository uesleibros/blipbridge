# The C ABI

`include/blipbridge/blipbridge.h` is the stable public interface. It is plain C:
no C++ types, no STL types, no exceptions and no allocation ownership cross it.
The COM Automation surface still exists, but it is now the research and
compatibility layer - nothing in normal use needs `regsvr32`, a ProgID,
`CreateObject`, or a COM add-in.

## Shape

```c
BB_Init / BB_Shutdown
BB_LoadTexture / BB_ApplyTexture / BB_ApplyTextureBatch
BB_ReleaseTexture / BB_ClearTextures / BB_GetTextureCount
BB_GetCapabilities / BB_GetLastError / BB_GetVersion / BB_GetVersionString
```

Every entry point returns `BB_Result` (`0` success, negative failure). Handles
are `uint64_t`, opaque, and never zero when valid. On Windows x64 there is one
calling convention, so exports are undecorated - which is what lets VBA bind by
plain name:

```text
BB_ApplyTexture  BB_ApplyTextureBatch  BB_ClearTextures  BB_GetCapabilities
BB_GetLastError  BB_GetTextureCount    BB_GetVersion     BB_GetVersionString
BB_Init          BB_LoadTexture        BB_ReleaseTexture BB_Shutdown
```

## Ownership

| Thing | Rule |
|---|---|
| Texture handle | owned by the library until released; never recycled |
| Image bytes | copied before `BB_LoadTexture` returns |
| Shape pointer | **borrowed for the call only**, never stored |
| Error / version strings | copied into the caller's buffer |

Passing `ObjPtr(shape)` is safe because the Shape is a live argument in the VBA
frame for the duration of the call, which keeps a reference alive. The native
side additionally validates before trusting it: the pointer must address
committed memory, must answer `QueryInterface` for `IDispatch`, and the object
must report an AutoShape or Freeform type. Never cache the value of `ObjPtr`.

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
    blipbridge.h                 stable C ABI, platform independent
         |
    src/abi/blipbridge_abi.cpp   thread affinity, error text, translation
         |
    src/backend/backend.hpp      the platform seam
         |
   +-----+---------------------------+
   |                                 |
windows_office_backend.cpp    unsupported_backend.cpp
   |                            (macOS and everything else)
experiments/exp_internal_blip
   the reverse-engineered PowerPoint implementation
```

Nothing above `backend.hpp` knows Office exists. The reverse-engineered code is
reached only through `bb::Backend`, and its exceptions are converted to
`BackendResult` there so none can reach the C ABI.

That directory is still named `experiments/` for historical reasons; it holds
the validated Windows backend and the research probes that produced it. Renaming
it would break a large number of evidence links in `docs/`, so it is deferred
rather than done hastily.

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
