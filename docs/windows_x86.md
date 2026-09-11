# Windows x86: what exists, and what it would take

**Status: the library builds and fails closed. There is no accelerated backend.**

Two statements that are easy to run together, and must not be:

| | |
|---|---|
| **x86 build** | Validated in CI. Debug and Release both compile, link, and produce a non-empty `BlipBridge-x86.dll` of the right machine type (`pei-i386`) exporting the full undecorated C ABI - `BB_ApplyTextureRange` and `BB_ApplyTextureIfChanged` included - and the ABI contract test runs against it. |
| **x86 Office runtime** | **Not validated.** No build of BlipBridge has ever been run inside a real 32-bit PowerPoint. The x86 backend refuses every texture call by design, so there is nothing there to accelerate; but "CI is green" is a statement about the build, never about the host. |

Nothing below claims otherwise.

A 32-bit build of BlipBridge loads, exports the full C ABI, reports its version
and capabilities, and answers every texture call with a specific refusal. It will
not fill anything. That is deliberate, and this page explains why it is not
simply a matter of recompiling.

## Why the x64 backend is not built into the x86 binary

It would compile. That is the problem.

The 64-bit backend is not portable code that happens to target Windows. It is a
reconstruction of a specific Office build's internals, and every part of it is
tied to the architecture:

| What it depends on | Why x86 differs |
|---|---|
| **Calling convention** | The x64 ABI passes the first four arguments in RCX/RDX/R8/R9 and returns a non-trivial class through a hidden pointer in RCX. 32-bit Office uses `thiscall` for member functions (`this` in ECX, the rest on the stack), `stdcall` for most exports, and returns structures differently. Every function-pointer typedef in `native_apply.hpp` encodes the x64 rules. |
| **Module RVAs** | Every private entry point is reached at a byte offset into `oart.dll`. A 32-bit `oart.dll` is a different binary with different offsets. None of the recorded RVAs mean anything there. |
| **Signature bytes** | Each entry point is byte-verified against the exact bytes its ABI was derived from. Those are x64 instruction encodings. They cannot match 32-bit code. |
| **Object layouts** | Offsets such as the FillFormat's `+0x58` control block and the record's `+0x2A0` stretch slot are pointer-size dependent. A structure of pointers packs differently at 4 bytes each. |
| **Pointer size in the record** | The property record is `0x4E8` bytes with pointer-sized slots. The 32-bit equivalent is a different size with different slot positions. |

If the backend were compiled as x86 it would resolve the same *numbers* against a
different binary, byte-verify against instructions that are not there, and - if a
signature check were ever relaxed - call into the middle of an unrelated
function. For a library that already found a way to terminate PowerPoint by
applying to a connector, that is not a risk worth taking for the sake of a second
file on the release page.

So the CMake build **excludes those sources entirely** on x86. There is no
`#ifdef` inside them that could be got wrong; the files are not compiled. The
x86 binary contains the public C ABI over `unsupported_backend.cpp`, and nothing
that knows what OART is.

## What the x86 package is for

It is genuinely useful, just not for speed:

* the C ABI can be exercised on 32-bit Office;
* the VBA wrapper's architecture detection and its 64-bit handle transport can be
  tested where `LongLong` behaves differently;
* the packaging, the checksums and the install flow can be validated;
* a caller gets a clear, specific message instead of a loader error.

## The x86 calling convention decision

One ABI decision was forced by 32-bit VBA and is already made.

`BB_CALL` is empty on x64, because the x64 ABI has exactly one convention. On x86
it is `__stdcall`, because **VBA's `Declare` can only call stdcall** and offers no
syntax to request anything else. A cdecl export would appear to bind and then
leave the stack unbalanced on every call - silent until it is not.

Exported names stay undecorated on both, so one `Declare` works on either
architecture. The 32-bit link passes `-Wl,--kill-at` to strip the `@N` suffix
stdcall would otherwise add; CI asserts the undecorated names on both
architectures, so losing that flag fails the build rather than the user.

## What validating an x86 backend would actually require

Not a port. A second reverse-engineering effort, repeating every step that
produced the x64 one, against 32-bit `PPCORE.DLL`, `OART.DLL` and `GFX.DLL`:

1. **Module identity** - locate and version-check the three modules in a 32-bit
   PowerPoint process.
2. **The wrapper** - confirm `Shape.Fill` is still a PPCORE delegating wrapper,
   decode its thunks, and read out the inner offset. The thunk *shape* differs:
   32-bit thunks are not the `48 8b 49 XX` sequence the x64 decoder matches.
3. **The FillFormat** - find the OART FillFormat vtable, its reference count
   offset, and the control block offset.
4. **The receiver** - re-derive the control block to receiver walk and the
   receiver's vtable.
5. **The private entry points** - locate the record constructor, the image
   sub-record, the transfer, the stretch holder, the transaction constructor,
   the destructors, and the apply slot. Derive each one's ABI, which on x86 means
   determining `thiscall` against `stdcall` per function and where the arguments
   actually sit.
6. **Record construction** - re-derive the record size and every slot offset for
   32-bit pointers.
7. **Ownership** - re-establish the reference accounting; the x64 conclusion that
   the operation owns exactly one counted reference must be re-proved, not
   assumed.
8. **Cached image creation** - confirm the GFX export names and their 32-bit
   signatures.
9. **Semantic eligibility** - rebuild the Shape compatibility matrix from
   scratch. **The x64 matrix must not be inherited.** That a Connector is fatal
   on x64 says nothing about what it is on x86, in either direction, and the
   whole point of that matrix is that it was measured rather than reasoned about.
10. **The full suite** - Undo, Redo, SaveAs, reopen, Shape deletion, slide
    deletion, presentation close, and the stress matrix, all repeated.
11. **Benchmarks** - measured separately. The x64 numbers do not transfer and
    must not be quoted for x86.

Only when all of that passes does x86 become a supported architecture, and only
then does its package stop carrying the warning it carries today.

## How to tell which PowerPoint you have

**File > Account > About PowerPoint.** The first line ends with `64-bit` or
`32-bit`. Download the matching package. If you get it wrong, `BlipBridge.bas`
says so in those words - it checks before it tries to load anything.
