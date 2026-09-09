# Reusable native textures

Office 16.0.14334.20848 x64, 2026-09-09.

```text
handle = LoadTexture(bytes)     decode once
ApplyTexture(shape, handle)     hot path: no decode, no file, no donor
ReleaseTexture(handle)
ClearTextures()
```

`LoadTexture` decodes bytes into a GFX cached image through the exported
`GEL::ICachedImage::Create`; `ApplyTexture` then costs only property-record
construction and the receiver call. Neither touches the filesystem, a donor
Shape, `PickUp`/`Apply`, or `Fill.UserPicture`.

## Ownership

A cached image is a **GFX-level** object: its creator takes an `IStream` and
nothing else - no presentation, no slide, no Shape. So a texture handle owns one
intrusive reference that is independent of any document. Three consequences,
each measured below:

* A handle stays valid across `Presentation.Close`, and applies fine in a
  different presentation afterwards.
* A handle must not outlive the GFX module. `ClearTextures` runs from the
  Engine's `OnDisconnection` **and** its destructor, both while Office is still
  loaded.
* Holding a texture needs no document; applying one needs a live Shape.

Handles never recycle, so a released handle is reported rather than silently
resolving to a different texture. Native handles start at `0x1000000`, disjoint
from the donor fallback's, so one `ApplyTexture` serves both backends and cannot
confuse them.

The **receiver is never cached**. It is re-resolved from `Shape.Fill` on every
single apply, because a Shape deleted through public COM still passes every
pointer and vtable check in the chain.

## Validation, and what is safe to cache

The first benchmark run made the hot path *slower* than `Fill.UserPicture`:
2.35 ms against 0.91 ms. The cause was the guards, not the work - every apply was
re-reading two module version resources off disk, decoding up to 48 vtable thunks
with a `VirtualQuery` apiece, and byte-verifying twelve function signatures.

What is now cached is exactly what is a property of a **loaded image** and cannot
change while it stays loaded:

| Cached | Key | Invalidated when |
|---|---|---|
| Module version check | `HMODULE` | any of oart/ppcore/gfx handle changes |
| Delegating-wrapper analysis | vtable address | as above |
| Byte-verified entry points | OART base address | OART loads at a different base |

What is **not** cached, and is re-checked on every apply: the FillFormat pointer,
its vtable, the control block, the receiver and the receiver's vtable. Those are
document-derived and can dangle; the module facts cannot.

That change alone took the hot path from 2.35 ms to 0.185 ms.

## Measured

### Benchmark, in process

Driving the measurement from PowerShell adds a cross-process COM round trip per
call, several times the cost of the operation itself, so the benchmark runs
inside PowerPoint (`tools/run_texture_benchmark.ps1` only asks for it). Same
Shape, same image, same document, 500 iterations each:

| | mean | median |
|---|---:|---:|
| `Fill.UserPicture(path)` | 0.6692 ms | 0.6337 ms |
| `ApplyTexture(handle)` | **0.1845 ms** | **0.1706 ms** |
| speed-up | **3.63x** | **3.71x** |

`LoadTexture` cost **0.0495 ms**, paid once per texture. It is recovered after
0.1 applies - that is, immediately.

The comparison is honest in the one direction that matters: `UserPicture` is
doing more work because Office re-decodes the image on every call, which is
precisely the work a texture handle removes.

### Lifetime matrix

`tools/test_native_texture.ps1`, all passing:

| Scenario | Result |
|---|---|
| One texture, 6 Shapes, 2 slides, incl. a Freeform | all filled, identity and geometry unchanged, nodes still editable |
| 1000 applies round-robin | 438.9 ms total; private bytes 99.7 -> 99.9 MB |
| Three textures at once (PNG, PNG, JPEG) | coexist and apply independently |
| Release out of order | the other handles keep working |
| Released handle | rejected |
| Unknown handle | rejected |
| Shape deletion | later applies unaffected |
| `Presentation.Close` | textures survive it |
| Apply in two other presentations | works, including the reopened deck |
| SaveAs and reopen | 6 picture fills, 0 Picture shapes |
| `ClearTextures` | count returns to zero |

`tools/test_native_texture_shutdown.ps1` covers the remaining boundary: five
textures still loaded when `Application.Quit` runs. PowerPoint exits cleanly, and
a fresh host starts with an empty store. That test runs a **control first** - a
host that never loaded a texture - because the first attempt failed only because
the harness itself still held COM references.

### Reference accounting

After 1006 applies of one texture over six Shapes the cached image's count sat at
60 and stopped there, with flat process memory. Sixty is what twenty undo entries
holding three references each would look like, and PowerPoint's undo history is
bounded by default - but that is a reading of the number, not a measurement of
it. The count is bounded and does not grow; the individual owners are still
unattributed.

## Why capabilities are still false

```text
MemoryImageToFill=False
CachedTextureApply=False
InternalBackend=False
```

Two things are outstanding:

1. **Undo/Redo is unproven.** `CommandBars.ExecuteMso('Undo')` returns `E_FAIL`
   in this automation harness after an ordinary `Fill.UserPicture` too, so the
   existing result says nothing either way. The real handler's wrapper
   `OART +0x8A13E0` sets up an action scope with `+0x1B8BE0`/`+0x1B8C60` that the
   native path skips, so the honest expectation is that a native apply is not
   undoable - but that has not been shown, and a harness that can actually drive
   Undo is needed before it can be.
2. **The reference count is bounded but unattributed.** Knowing it stops at 60 is
   not the same as knowing who holds those sixty.

Everything else on the productization list has been measured and passes.
