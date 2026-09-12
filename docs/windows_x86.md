# Windows x86: the portable backend

**The full public API is implemented for the x86 build through the portable
backend. Portable backend behaviour is validated against real PowerPoint using
forced-portable x64 testing. Runtime validation inside real 32-bit PowerPoint
remains outstanding.**

Those are three different statements and this page keeps them apart. Blurring
them is the easy mistake, so here they are as four separate facts, each with what
backs it:

| | | Evidence |
|---|---|---|
| **x86 implementation** | yes | The portable backend implements the whole `Backend` interface; x86 compiles it instead of the accelerated one. |
| **x86 compilation and CI** | yes | Debug and Release build in CI and pass the full contract suite, including the COM contract in a 32-bit process. |
| **Portable backend behaviour against real PowerPoint** | yes | `tools/test_portable_matrix.ps1` runs the whole public surface against the portable backend inside a real PowerPoint, via `BB_FORCE_PORTABLE_BACKEND=ON`. |
| **Real 32-bit PowerPoint runtime** | **not yet validated** | No build of BlipBridge has been executed inside a 32-bit `POWERPNT.EXE`. Office does not install both architectures side by side, and the development machine has 64-bit Office. |

The fourth row is the one that matters when deciding whether to depend on this.
The gap it names is the compiler, not the logic: the portable backend is
architecture-neutral by construction, and forced-portable testing exercises every
line of it that could be wrong in a real host. That is a reason to expect it to
work on 32-bit PowerPoint. It is not evidence that it does.

## What x86 does, and how

Everything the public API promises: textures from encoded bytes, raw BGRA or a
file; `BB_ApplyTexture`, `BB_ApplyTextureBatch`, `BB_ApplyTextureRange`,
`BB_ApplyTextureIfChanged` and its skip cache; `BB_ApplyPicture`; and the whole
image pipeline - decode, crop, orient, resize and quad warp.

It does that through documented Office Automation - `Shape.Fill.UserPicture` and
`ShapeRange.Fill` - rather than through Office internals. **`BB_CAP_NATIVE_BACKEND`
is not set on x86**, so the capability mask there is `0x03FE` rather than
`0x03FF`, and a caller who measures will find it slower. See
[benchmarks.md](benchmarks.md) for how much.

## What changed, and what did not

x86 previously had no backend at all. The library loaded, reported its version,
and refused every texture call with a specific reason. That is over: there is now
a **portable backend** (`src/backend/portable_office_backend.cpp`) that
implements the same `Backend` interface over Office's own API.

What has *not* changed is the position on the accelerated backend. None of the
reverse-engineered OART/GFX work is compiled on x86, for exactly the reasons
below. Two backends exist; x86 gets the portable one.

## Why the x64 backend is still not built into the x86 binary

It would compile. That is the problem.

The accelerated backend is not portable code that happens to target Windows. It
is a reconstruction of a specific Office build's internals, and every part of it
is tied to the architecture:

| What it depends on | Why x86 differs |
|---|---|
| **Calling convention** | The x64 ABI passes the first four arguments in RCX/RDX/R8/R9 and returns a non-trivial class through a hidden pointer in RCX. 32-bit Office uses `thiscall` for member functions (`this` in ECX, the rest on the stack), `stdcall` for most exports, and returns structures differently. Every function-pointer typedef in `native_apply.hpp` encodes the x64 rules. |
| **Module RVAs** | Every private entry point is reached at a byte offset into `oart.dll`. A 32-bit `oart.dll` is a different binary with different offsets. None of the recorded RVAs mean anything there. |
| **Signature bytes** | Each entry point is byte-verified against the exact bytes its ABI was derived from. Those are x64 instruction encodings. They cannot match 32-bit code. |
| **Object layouts** | Offsets such as the FillFormat's `+0x58` control block and the record's `+0x2A0` stretch slot are pointer-size dependent. A structure of pointers packs differently at 4 bytes each. |
| **Pointer size in the record** | The property record is `0x4E8` bytes with pointer-sized slots. The 32-bit equivalent is a different size with different slot positions. |

If it were compiled as x86 it would resolve the same *numbers* against a
different binary, byte-verify against instructions that are not there, and - if a
signature check were ever relaxed - call into the middle of an unrelated
function. For a library that already found a way to terminate PowerPoint by
applying to a connector, that is not a risk worth taking.

So the CMake build **excludes those sources entirely** on x86. There is no
`#ifdef` inside them that could be got wrong; the files are not compiled. The
x86 binary contains the public C ABI, the image pipeline, and a backend that
knows nothing about OART.

## How the portable backend works

Every route ends at `Fill.UserPicture`, which takes a **path**. So a texture here
holds its decoded BGRA pixels and writes them to a temporary PNG the first time
it is applied - once per texture, not once per apply, because the file does not
change. The file lives as long as the texture and is deleted with it, in
`%TEMP%\BlipBridge-<pid>\`.

That gives the two resources opposite shapes, which is worth stating plainly:

| | accelerated (x64) | portable (x86) |
|---|---|---|
| pixels retained by BlipBridge | no | yes, BGRA32 |
| first apply | one private call | encode + write, then Office reads the file |
| repeat apply | one private call | Office reads the same file again |
| what Office stores | one shared cached image | an embedded copy per fill |
| an apply that changes nothing | skipped by the cache | skipped by the same cache |
| undo entries for a range apply | one | one |

The last two rows matter: the skip cache, the Shape identity key and the Shape
class policy are **the same code on both backends**, not a second
implementation. They decide which Shapes may be filled, which may be cached, and
when an apply may be skipped - and a divergence there would surface to a caller
as "it behaves differently on 32-bit" with nothing to point at.

### Where the two do differ

Three differences, all deliberate, all measured:

1. **Speed.** There is no private transaction to bypass. An apply costs what
   Office charges for an apply, plus one PNG encode the first time a texture is
   used. The x64 benchmark numbers do not transfer and are not quoted for x86.

2. **`BB_CAP_NATIVE_BACKEND` is clear.** The capability mask on x86 inside
   PowerPoint is `0x03FE` - every feature bit set, the accelerated bit not. A
   caller who benchmarks and finds it slower can discover why by asking, rather
   than by guessing.

3. **Undo grouping.** The accelerated apply opens its own transaction, so it is
   always its own undo entry. `Fill.UserPicture` joins whatever undo entry Office
   currently has open, so a caller who creates a Shape and immediately fills it
   may find one Undo reverts both. This is ordinary Office behaviour - it is what
   plain VBA `Fill.UserPicture` has always done - and `Application.StartNewUndoEntry`
   is the way to force a boundary. A range apply is still exactly one entry on
   both.

### Temporary files, and their lifecycle

The portable backend writes files. `Fill.UserPicture` takes a path, so a texture
built from raw pixels - a resample, a crop, a quad warp - has to become one. This
is stated plainly rather than buried, because it is a real difference: the
accelerated backend hands Office an in-memory image and writes nothing.

The rules:

* **Where.** `%TEMP%\BlipBridge-<pid>\`. Per-process, so two PowerPoint
  instances with the add-in loaded cannot delete each other's files, and no
  locking is needed to guarantee it.
* **What.** One PNG per texture, named from the texture's process-unique image
  id, so two textures never collide.
* **When.** On the texture's **first apply**, not at load. A texture that is
  created and never applied never touches the disk. Encoding is the expensive
  part and the pixels do not change, so a texture applied to fifty Shapes encodes
  once.
* **Until when.** The file lives as long as the texture and is deleted with it -
  on `BB_ReleaseTexture`, on `BB_ClearTextures`, or when the last owner drops it.
* **The directory.** `BB_ClearTextures` removes it once empty. It is recreated
  lazily on the next write, so a cleared library keeps working without being
  re-initialised. That was a real bug once - clearing left every later apply
  failing with "the system cannot find the path specified" - and
  `tools/test_portable_matrix.ps1` asserts the load-apply-clear-load-apply cycle
  because of it.
* **When cleanup fails.** Best effort. A file that will not delete is a stray
  temporary, not a reason to fail a destructor or an apply; the directory is only
  ever removed when empty, so a file that outlived its texture is left where a
  human can find it rather than taken down with something else.
* **What is never touched.** Only files this process wrote, only inside its own
  directory.

A caller who needs to know whether anything is on disk can ask: the research
surface's texture report includes `encoded=1` once a texture's file exists.

## What the portable backend refuses, and why it refuses more than it must

`BB_ApplyTextureRange` refuses a range containing a Table, even though
`Fill.UserPicture` would accept one and this backend could therefore fill it.

That is a deliberate trade. The range API has to mean one thing: a range that
filled on 32-bit and was refused on 64-bit would be a difference a caller
discovers in production, from a document that came out wrong, with nothing in the
API to explain it. Being uniformly stricter costs a capability nobody was
promised; being non-uniformly permissive costs trust in the whole surface.

`BB_ApplyPicture` is the entry point that accepts fallback-only classes such as
Tables - on both backends, as it always has.

## How it was validated

The portable backend is architecture-neutral by construction: the only
32-bit-specific thing about it is the compiler. So it is validated two ways, and
neither alone would be enough.

**In a real PowerPoint.** Configuring with `-DBB_FORCE_PORTABLE_BACKEND=ON`
builds the portable backend in place of the accelerated one, including into the
research COM surface the Office harnesses drive. The harnesses then exercise the
portable backend inside a real PowerPoint, through the real C ABI. These suites
pass against it:

| suite | what it covers |
|---|---|
| `test_portable_matrix` | the whole public surface, end to end - see below |
| `test_image_api` | the ABI 5 image surface, quad warp applied to a Shape, save and reopen |
| `test_apply_if_changed` | the skip cache, staleness, refusals, WordArt and Freeform, save and reopen |
| `test_range_apply` | range fills, all-or-nothing refusal, one undo entry, mixed classes, groups |
| `test_picture_cache` | `BB_ApplyPicture`, its caches, and a file rewritten on disk |
| `test_shape_lifecycle_cache` | delete, undo, redo, groups, and cache identity across all of it |
| `test_cache_ownership` | the two owners, and released handles staying stale forever |
| `test_semantic_guards` | Connector, Line and WordArt |
| `test_release_stress` | 2000 applies, leak behaviour, and the store's own accounting |

`test_portable_matrix.ps1` is the one written for this release, and it is
deliberately **backend-aware rather than portable-only**: it reads which backend
is loaded, asserts the capability mask that backend owes, and then asserts
identical behaviour for everything else. Running it against both builds is what
turns "the same API either way" from a design intention into a measured claim.

Where a result is visible it samples the rendered Shape rather than asking a
cache what it believes. `Fill.Type` reports that a picture is present and never
which one, and the failure worth guarding against is a cache that shows the wrong
picture confidently.

The suites that do **not** run against it are the ones that instrument Office
internals - receiver layouts, transaction splitting, GFX image lifetimes. There
is nothing there for them to look at, and their Automation ids say so rather than
reporting "unknown member".

**As an x86 binary.** CI builds Debug and Release for x86 and runs the full
contract suite, which now includes the PNG encoder the portable backend depends
on. The 32-bit build produces `BlipBridge-x86.dll` of the right machine type
exporting the full undecorated C ABI.

### What is still not claimed

**No build of BlipBridge has been executed inside a real 32-bit `POWERPNT.EXE`.**
The development machine has 64-bit Office, and Office does not install both
architectures side by side.

So the claim is worded exactly this way, and no more strongly:

> The full public API is implemented for the x86 build through the portable
> backend. Portable backend behaviour is validated against real PowerPoint using
> forced-portable x64 testing. Runtime validation inside real 32-bit PowerPoint
> remains outstanding.

That is a much stronger position than "it compiles", and it is not the same thing
as "validated on 32-bit Office". Anyone deciding whether to depend on the x86
package should read it as: the logic is tested, the composition with a 32-bit
host is not.

## The x86 calling convention decision

`BB_CALL` is empty on x64, because the x64 ABI has exactly one convention. On x86
it is `__stdcall`, because **VBA's `Declare` can only call stdcall** and offers no
syntax to request anything else. A cdecl export would appear to bind and then
leave the stack unbalanced on every call - silent until it is not.

Exported names stay undecorated on both, so one `Declare` works on either
architecture. The 32-bit link passes `-Wl,--kill-at` to strip the `@N` suffix
stdcall would otherwise add; CI asserts the undecorated names on both
architectures, so losing that flag fails the build rather than the user.

## What an accelerated x86 backend would still require

Unchanged by any of the above. Not a port - a second reverse-engineering effort,
repeating every step that produced the x64 one, against 32-bit `PPCORE.DLL`,
`OART.DLL` and `GFX.DLL`:

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

All of that needs a 32-bit Office installation to work against, which is the
first thing it needs and the thing that is missing. Until then x86 has a backend
that works and says plainly that it is not the fast one.

## How to tell which PowerPoint you have

**File > Account > About PowerPoint.** The first line ends with `64-bit` or
`32-bit`. Download the matching package. If you get it wrong, `BlipBridge.bas`
says so in those words - it checks before it tries to load anything.
