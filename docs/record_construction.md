# How the fill property record is actually built

Observed on Office 16.0.14334.20848 x64, 2026-09-09. Static disassembly of
`OART +0x89C860`, the `Fill.UserPicture` handler, corroborated by the live record
dumps in `fill_transaction.md`.

The question this answers: of the roughly 0x4E0 bytes of the property record, how
much would a native apply actually have to produce. The answer is four slots -
and Office's own constructors produce the rest.

## The handler's whole record build

Between the guard checks and the commit, the handler does only this:

```c
// OART +0x89C9AA .. +0x89CA48, record == rbp-0x80
+0x89CEF8();                              // precondition check
+0x14110(record);                         // property record constructor
+0x14F7B0(record);                        // mark every property slot unset
*(uint32_t*)(record + 0x04) = 3;          // payload of the slot at +0x00
*(uint32_t*)(record + 0x00) |= ...;       // (value & ~6) | 1  -> slot is set

+0x14580(loaded);                         // image sub-record constructor
+0x950070(loaded, filename, true, false); // load the file into it

+0x22BCF4(record + 0x88, loaded);         // move the image into the record slot
+0x15AC70(record + 0x2A0, handlerArg3);   // set the 16-byte-value slot
*(uint8_t*) (record + 0x4DC) = 1;
*(uint32_t*)(record + 0x4D8) |= ...;      // (value & ~6) | 1  -> slot is set

Transaction t;
+0x48870(&t, record, 0, handler->byte_at_0x60, 0xA042008E);
receiver->vtable[0x78](receiver, &t);
```

Nothing else touches the record. Everything not listed is whatever `+0x14110`
and `+0x14F7B0` leave behind.

## The record is a list of tagged slots

`OART +0x14F7B0` is a 60-byte leaf that ANDs `0xFFFFFFFA` into a fixed list of
dwords. That list is the record's slot table:

```text
+0x000  +0x008  +0x028  +0x030  +0x050  +0x068  +0x07C
+0x088  +0x288  +0x2A0  +0x2B8  +0x4A8  +0x4B0
```

Each slot begins with a discriminator dword. Two idioms appear throughout:

* `value &= 0xFFFFFFFA` - the slot is unset.
* `value = (value & ~6) | 1` - the slot now holds a value.

`+0x088` is the image slot, which is why the transfer at `+0x22BCF4` is called
with `record + 0x88` and why the live dump reads `record+0x88 = 1` afterwards.
`OART +0x15AC70` is the same idiom as a helper: assign through `+0x249900` into
`slot + 8`, then stamp the discriminator.

`+0x4D8` is not in the slot table, so it is a separate tagged field whose payload
is the byte at `+0x4DC`.

## Which fields are which

| Offset | Role | Kind |
|---|---|---|
| +0x000 / +0x004 | slot with payload `3` | constant for a picture fill |
| +0x088 | image slot discriminator | image-specific |
| +0x090 | image sub-record, cached GFX image at its +0xF0 | image-specific, owned |
| +0x2A0 | slot holding a counted 16-byte value | see below |
| +0x4D8 / +0x4DC | tagged flag, payload `1` | constant |
| everything else | left unset by +0x14F7B0 | default |

The value in slot `+0x2A0` comes from the handler's third argument, which
`OART +0x8A13E0` builds locally with `+0x158C40` from **sixteen zero bytes**;
`+0x158C40` allocates 0x10 bytes, copies them, and wraps them with
`OART +0xA45DB0`. The live transaction dump agrees: `record+0x2A8` is a heap
pointer and `record+0x2B0` is `oart.dll+0xA45DB0`. Sixteen bytes wrapped like
this, in a pipeline whose image creator takes an `Ofc::MD4UID*`, reads as a
content UID that PowerPoint leaves zero. That is a reading of the evidence, not
a recovered name.

Ownership: the record owns one counted reference to the cached image, taken by
the copy path `+0x13360 -> +0xA0870 -> +0x9848C` and released by the destructor
path `+0xAF80 -> +0x3D870 -> +0x3DBAB`. Everything else in the record is either
inline or held through the same counted-slot idiom.

## Office's own constructors would do the work

The user-visible question was whether a synthetic record must be fabricated byte
by byte. It must not. The constructors exist and are the ones the handler itself
uses:

| Function | Role |
|---|---|
| `OART +0x14110` | property record constructor |
| `OART +0x14F7B0` | mark all slots unset |
| `OART +0x14580` | image sub-record constructor |
| `OART +0x8F94C` | install a cached image into an image sub-record |
| `OART +0x22BCF4` | move an image sub-record into the record's image slot |
| `OART +0x15AC70` | set a counted-value slot |
| `OART +0x48870` | transaction constructor |
| `OART +0xAF80` | property record destructor |

`OART +0x950070` is the file loader and is the one link a native path would
replace. Its role is only to produce a cached image and hand it to `+0x8F94C`,
which stores it at record+0xF0 and derives metadata through the image's own
vtable slot +0x20.

## The cached image creator is an export, not a private offset

`GFX.DLL` exports the stream-based creator by name, ordinal 236 at RVA 0x7680:

```text
?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@
    AEAV?$TCntPtr@UIImage@GEL@@@4@ PEAUIStream@@
    W4IStreamCopyInstruction@12@ PEBVMD4UID@4@ _N @Z
```

Demangled, and confirmed against its prologue:

```c
// RCX is the hidden sret pointer: GFX +0x7680 begins `mov %rcx,0x8(%rsp)`
// and returns RAX = RCX, the MSVC x64 convention for a non-trivial return.
Ofc::TCntPtr<GEL::ICachedImage>
GEL::ICachedImage::Create(
    Ofc::TCntPtr<GEL::IImage>&              image,        // RDX
    IStream*                                stream,       // R8
    GEL::ICachedImage::IStreamCopyInstruction instruction, // R9D
    const Ofc::MD4UID*                      uid,          // [rsp+0x28]
    bool                                    flag);        // [rsp+0x30]
```

It takes a **standard `IStream*`**, which is exactly what a byte array can be
turned into without touching the filesystem. The prologue also settles ownership:
the worker's result is *moved* into the return storage (`*rax` is read then
zeroed) rather than AddRef'd, so the returned pointer carries one reference the
caller must release through vtable slot +8.

Because this is an export, a native loader resolves it with `GetProcAddress`,
not with a hard-coded RVA. Every other function in the table above is a private
offset and would need its own validation and guard.

Office's own call site, `OART +0x8F54E`, supplies these arguments:

```text
rcx        = &TCntPtr<ICachedImage>   (return storage)
rdx        = &TCntPtr<IImage>         (in/out, empty on entry)
r8         = IStream*
r9d        = 0                        (IStreamCopyInstruction)
[rsp+0x20] = const MD4UID*            (a 16-byte local)
[rsp+0x28] = false                    (bool)
```

and hands the result straight to `OART +0x8F94C(record, &cached)`.

## This half is implemented and validated

`experiments/exp_internal_blip/cached_image_load.cpp` calls exactly that export
and releases what it creates. It touches no Shape, slide, presentation or
property record, so it cannot damage a document. `tools/test_cached_image_load.ps1`
exercises it.

Measured, and matching the static prediction in every field:

```text
texture_64_1.png  (11822 b): cachedCount=1 image=... imageCount=2 imageVtableMatches=1
texture_64_1.png  (11822 b): cachedCount=1 image=... imageCount=2 imageVtableMatches=1
texture_64_1.png  (11822 b): cachedCount=1 image=... imageCount=2 imageVtableMatches=1
texture_256_1.png (203682 b): cachedCount=1 image=... imageCount=2 imageVtableMatches=1
texture_64_0.jpg   (2265 b): cachedCount=1 image=... imageCount=2 imageVtableMatches=1
```

The cached image's vtable is `GFX +0x409DC0` and the image's is `GFX +0x4055C8`,
as the debugger observed. The counts 1 and 2 are exactly the factory-return row
in `resource_lifetime.md`, now reproduced from our own call rather than by
watching Office make it. Invalid bytes are rejected without a crash, and an
ordinary `Fill.UserPicture` still works afterwards.

PNG and JPEG bytes both decode with **no file on disk and no `UserPicture`**.
This is the "load once" half of the target architecture working.

Two things it is not. It is not a texture handle: both references are released
before the call returns, because nothing is yet safe to retain across a document
lifetime. And it is not a performance result: no apply exists to measure.

The identical PNG decoded three times reported the same *image* address each
time, but each had been released before the next call, so that is address reuse
after free and is not evidence that GFX caches by content.

## Remaining work for the apply half

## Measured: Office does not reuse cached images

The receiver experiment filled the same Shape twice with the same file and two
other Shapes with the same file again. **The cached GFX image differed on all
four calls.** Office decodes and builds a new cached image on every
`Fill.UserPicture`, even for an identical path.

That is the headroom the project is aiming at: a reusable texture handle would
skip work that Office currently repeats per call. It is not yet a performance
claim - nothing has been measured against a native apply, because no native
apply exists.

Per the project's own gating rules, all of these remain open before any of the
private OART functions above is called:

* The exact size of the property record and of the image sub-record.
* What `+0x14110` and `+0x14580` leave behind, and what their destructors require.
* Dynamic validation of each private function's ABI, rather than reading it off
  the disassembly.
* Whether a record built outside the handler survives undo, save and reopen.

No private Office *offset* is called anywhere in the current code. What ships is
one exported GFX function, resolved by name and validated above, plus the
guarded read-only lookups in `experiments/exp_internal_blip/receiver_inspect.cpp`.
Every OART function in the table above remains uncalled.

Evidence: `docs/evidence/receiver_identity.txt`,
`docs/evidence/gfx_stream_exports.txt`, `docs/fill_transaction.md`.
