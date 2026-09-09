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

## Lifetime matrix


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

## Undo and Redo

**Native applies are undoable and redoable.** An earlier run concluded nothing
because the control failed too, and the reason was the harness: those runs used
`Presentations.Add(0)`, a presentation with **no window**, and PowerPoint's Undo
acts on a document window.

`tools/test_undo_harness.ps1` fixes that and refuses to judge the native path
until a driver demonstrably undoes an ordinary `Fill.UserPicture`:

```text
control/ExecuteMso : UNDO WORKS (UserPicture reverted)
control/redo       : undo=True redo=True
native/undo        : Fill.Type 1 -> 6 -> 1   (undo reverted: True)
native/redo        : Fill.Type after redo = 6 (restored: True)
```

So the earlier expectation - that skipping the wrapper's action scope would make
the native apply non-undoable - was **wrong**. The apply goes through Office's own
transaction and receiver, and the undo entry comes with it.

## Reference ownership, attributed

Bounded memory said nothing was leaking; it did not say who held what.
`tools/test_refcount_attribution.ps1` walks one texture through a controlled
sequence, changing one thing at a time:

| Step | count | delta |
|---|---:|---:|
| `LoadTexture` | 1 | +1 |
| apply to Shape A | 5 | +4 |
| apply to Shape B, same slide | 9 | +4 |
| apply to Shape C, other slide | 12 | +3 |
| re-apply to Shape A | 13 | +1 |
| Undo | 14 | +1 |
| Redo | 15 | +1 |
| flush undo with 40 unrelated entries | 12 | -3 |
| delete Shape B | 11 | -1 |
| flush undo again | 5 | -6 |
| SaveAs | 5 | 0 |
| **`Presentation.Close`** | **1** | **-4** |

Closing the document returns the count to exactly **1** - the reference the
handle itself owns. Everything above 1 is document-owned and released with the
document. Reproduced three times: the original deck, a brand-new presentation
(+4 then back to 1), and the stress run's three concurrent presentations.

That also answers why identical applies produce different deltas. The cost of an
apply depends on **what the Shape's previous fill was** and **what is already in
the undo history**: a first fill over a solid Shape costs +4, while re-applying
the same image over itself costs +1 because the outgoing fill's reference moves
into the undo entry rather than adding a new one. The earlier +3/+2/+3 variation
is exactly this, not an inconsistency.

Owners identified: **handle-owned** (1, ours), **shape/fill-owned** (released on
delete), **undo-owned** (released by flushing history), and the remainder
**document-owned** (released on close). No reference outlives its document.

## Stress

`tools/test_native_texture_stress.ps1`:

| Scenario | Result |
|---|---|
| 10,000 applies, one Shape | 4236 ms, private 114.2 -> 122 MB, **0 extra creations** |
| 2,000 alternating applies | private flat at 122.2 MB, counts bounded (24 / 21) |
| 120 Shapes over 3 slides, one texture | all filled, count 364 (~3 per Shape, matching the attribution) |
| 20 Shape deletions, then a slide deletion | applies still work |
| 5 rounds x 100 load/release | private flat, handles back to 2, creations exactly 502 |
| double release, stale handle | both rejected |
| 3 concurrent presentations | one texture serves all |
| all presentations closed | **count back to 1** |

`creations` equalling the number of `LoadTexture` calls is what makes the reuse
claim concrete rather than rhetorical.

## Benchmark

In process, 1000 iterations, same Shape and image:

| | mean | median | p95 | p99 | max |
|---|---:|---:|---:|---:|---:|
| `Fill.UserPicture` | 0.6588 ms | 0.6189 ms | 0.8348 ms | 1.6002 ms | 2.6318 ms |
| `ApplyTexture`, same texture | **0.1860** | **0.1756** | **0.2260** | **0.3361** | 1.5537 |
| `ApplyTexture`, alternating two | 0.1857 | 0.1774 | 0.2309 | 0.2907 | 0.4037 |
| `LoadTexture` (50 samples) | 0.0231 | 0.0226 | 0.0235 | 0.0393 | 0.0393 |
| `ReleaseTexture` (50 samples) | 0.0001 | 0.0001 | 0.0001 | 0.0007 | 0.0007 |

Mean speed-up **3.54x**, median **3.52x**, alternating **3.55x**. Alternating
between two textures costs the same as repeating one, so there is no penalty for
switching. The tail is the bigger story: p99 **0.336 ms** against
`UserPicture`'s **1.600 ms**.

`LoadTexture` is recovered after 0.05 of one apply.

## Status

| | |
|---|---|
| Native cached apply | **proven** |
| File-free decode (no temporary image) | **proven** |
| Reuse across Shapes, slides, presentations | **proven** |
| Save / reopen | **proven** |
| Performance win | **proven** |
| Undo / Redo | **proven** |
| Reference ownership | **fully attributed** |

Capabilities are computed at runtime from the live modules; see
`capabilities.md` for what each flag claims and the evidence behind it.
