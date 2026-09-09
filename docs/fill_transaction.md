# Loaded image to OART transaction and the fill operation

Observed on OART 16.0.14334.20848 x64 and GFX 16.0.14334.20848 x64, 2026-09-09.
The experiment is `experiments/exp_internal_blip/probe_fill_transaction.py`,
driven by `tools/run_office_probe.ps1`. It sets validated breakpoints during
an ordinary `Fill.UserPicture` call; it does not invoke internal methods, write
inferior memory, or retain pointers after detach.

Every RVA below is verified against recorded instruction bytes before the probe
attaches. Ignore nearest-export names in raw GDB stacks; use these RVAs.

## Resource propagation

The same cached-image address was observed in each of these live records:

| Stage | Offset | Evidence |
|---|---:|---|
| Loaded image record | +0xF0 | Input RDX at OART +0x22BCF4 |
| Destination wrapper | +0xF8 | Wrapper embeds copied record at +8 |
| Transaction | +0x198 | Input RDX at OART +0x89CA6B |
| Operation | +0x1E8 | Input RDX at OART +0x21BC06 |

Before transfer, the wrapper discriminator at +0 was 2. OART +0x22BCF4 calls
+0xA0870 with destination+8 and the loaded source, then changes discriminator
bits with `(value & ~6) | 1`. This is an observed property-record transfer, not
a standalone image-to-Shape API.

The four offsets are one layout, not four unrelated ones. The property record
embeds an **image sub-record at record+0x90**, and that sub-record holds the
cached GFX pointer at its own **+0xF0**. So:

```text
property record + 0x90 (image sub-record) + 0xF0 (cached image)
transaction + 0x18 (property record)  -> cached image at transaction + 0x198
operation   + 0x68 (property record)  -> cached image at operation   + 0x1E8
```

The standalone "loaded record" observed at OART +0x22BCF4 is an instance of the
image sub-record type, which is why its cached pointer sits at +0xF0 directly.

## Resolved dispatch path

```text
UserPicture loader returns                   OART +0x89C9F6
loaded-record transfer                       OART +0x22BCF4
transaction construction                     OART +0x48870
receiver resolved from handler state +0x58   OART +0x63EA0
receiver vtable +0x78 call                   OART +0x89CA6B
    receiver vtable                          OART +0x9F6658
    resolved forwarding method               OART +0x21DEC0
    forwards unchanged to receiver slot +0x50
        driver                               OART +0x21BB20
        slot +0x60 call: build operation     OART +0x21BBE8 -> +0x2290F0
        slot +0x58 call: apply operation     OART +0x21BC06 -> +0x1B88B0
        operation vtable +0xA8, EDX=1        OART +0x21BC26 -> +0x66C80
returns to UserPicture                       OART +0x89CA71
```

`OART +0x21BB20` is a plain build / apply / delete driver:

```c
Operation* operation = nullptr;
if (receiver->vtable_0x60(receiver, transaction, &operation) && operation) {
    receiver->vtable_0x58(receiver, operation);   // apply
    operation->vtable_0xA8(operation, 1);         // destroy and free
}
```

## Operation construction

`OART +0x2290F0` performs RTTI-style type checks on the transaction. In the
observed run none of its own inline construction branches matched, so it fell
through to `OART +0x21BC50`, which reached the generic factory:

```text
OART +0x2290F0  +0x2291FE  ->  OART +0x21BC50
OART +0x21BC50  +0x21BCA4  ->  OART +0x2F1E00   (cold continuation)
OART +0x2F1E00  +0x2F1E1C  ->  OART +0xF740     (transaction type dispatch)
OART +0xF740    +0xF7F1    ->  OART +0xE7F0     (operation factory)
OART +0xE7F0    +0xEF3B    ->  OART +0x10244    (operation constructor)
```

The live backtrace at the constructor call matched this chain exactly in both
runs (+0xEF3B, +0xF7F6, +0x2F1E21, +0x229203, +0x21BBEE, +0x89CA71).

`OART +0xE7F0` allocates the operation from the Office allocator singleton at
`[OART +0xD40050]` through its vtable slot 0, then calls the constructor:

```c
// OART +0xEEFC .. +0xEF40
void* storage = allocator->vtable_0(allocator, 0x570);
OperationCtor(storage,
              transaction + 0x18,                     // RCX/RDX
              *(uint32_t*)(transaction + 0x500),      // R8D  flags
              *(uint8_t*) (transaction + 0x508),      // R9B  bool
              *(uint32_t*)(transaction + 0x504));     // stack, identifier
```

Observed values, identical in two independent runs:

| Input | Value |
|---|---|
| Allocation size | 0x570 |
| Source | transaction + 0x18 (verified equal at run time) |
| Flags (transaction +0x500) | 0 |
| Bool (transaction +0x508) | 0 |
| Identifier (transaction +0x504) | 0xA042008E |
| Resulting vtable | OART +0x9E4BD0 |

`OART +0x10244` sets the two vptrs (+0x9E4BD0 at +0, +0x9E47C0 at +8), takes a
process-wide sequence number from `[OART +0xD40040]` into +0x10, initialises a
sub-object at +0x38, and **copies the source property record into +0x68** via
+0x13360 / +0x10CD0 / +0x11E60. It stores the identifier at +0x568 and the bool
at +0x56C. It receives no Shape pointer, no document pointer, and no context.

## Where Shape and document identity enter

Identity does **not** enter through the operation. The operation is built purely
from the transaction's property record plus three scalars. Identity is carried
by the receiver, which `OART +0x63EA0` resolves from handler state +0x58:

```text
receiver +0x00  vtable OART +0x9F6658
receiver +0x08  PPCORE object, vtable ppcore.dll+0x1396DB8 (shared per slide)
receiver +0x10  smart pointer holding the shared null singleton OART +0x9E5690
receiver +0x18  vtable OART +0x9F64D0   -- the context sub-object
```

`OART +0x1B88B0` applies the operation against that receiver:

```c
operation->word_at_0x28 = 0x100;
if (!receiver->vtable_0x48(receiver)) return;
context = receiver->vtable_0xF0(receiver);        // returns receiver + 0x18
extra   = receiver->vtable_0xF8(receiver, &tmp);
OperationApply(operation, context, receiver + 0x10, &nullHolder, extra);
                                                  // OART +0x1B8F50
```

Live registers matched: RCX = operation, RDX = receiver+0x18, R8 = receiver+0x10.

### The receiver is per Shape

`receiver_lookup.md` is the authority for this; the summary is that the receiver
is stable per Shape, `receiver+0x8` is stable per slide, and nothing in the
property record identifies the Shape.

`experiments/exp_internal_blip/probe_receiver_identity.py`, driven by
`tools/prepare_receiver_identity.ps1`, fills the same PNG in the order A, A,
B (same slide as A), C (other slide). The repeat on A is what separates a
per-call object from a per-Shape one; the first version of this experiment used
three different Shapes and could not.

Measured, and reproduced:

| Value | Classification |
|---|---|
| handler state | stable per Shape |
| receiver token (handler +0x58) | stable per Shape |
| receiver | stable per Shape |
| receiver +0x8 | stable per slide |
| receiver +0x20 | stable per Shape (allocation counter, see below) |
| cached GFX image | differs per call |
| record +0x40 | differs per call |

So Shape identity is carried by the receiver. `receiver+0x8` points to a PPCORE
object whose vtable is `ppcore.dll+0x1396DB8`; it is the *same* object for the
two Shapes on one slide and a different one for the Shape on the other slide,
which is consistent with a slide-level PPCORE container. That identification is
inferred from the sharing pattern and a stable vtable, not from a symbol.

`receiver+0x20` was described here as a per-Shape index. That is **corrected** in
`receiver_lookup.md`: the receiver constructor `OART +0x223950` writes it once
from a process-global counter at `[OART +0xD40038]`, so it is an allocation
sequence number, not an identifier of anything in the document.
`receiver+0x28..+0x48` holds inline UTF-16 fragments with `receiver+0x38`
pointing at `receiver+0x48`, which reads as a small-buffer string rather than a
reference.

The property record prefix below the image sub-record is identical across every
call except `record+0x40`, a non-polymorphic heap pointer (first qword zero).
With the repeat on shape A added, that pointer is neither Shape-stable nor
slide-stable, and two runs disagree in a way only heap-address reuse explains.
It was the last candidate for Shape identity inside the record and it is not one.
It remains unidentified, but it is per call.

**Consequence for a native backend.** Neither the operation nor the property
record names the target Shape, so a hypothetical `ApplyTexture(shape, handle)`
cannot be built by constructing a record alone; it also needs the per-Shape OART
receiver. That receiver **is** reachable from the public `Shape.Fill` through
four guarded pointer loads - see `receiver_lookup.md`, which also supersedes the
claim made here that the receiver's Shape field was unidentified. The record
recipe is in `record_construction.md`.

## The property record and the transaction

The record and the transaction are both plainer than the dispatch around them.
All of this is static disassembly of `OART +0x89C860`, corroborated by the live
dumps.

The whole handler builds one property record in its own frame at `rbp-0x80`, and
the pieces observed separately are all inside it:

```text
record + 0x40   unidentified non-polymorphic pointer
record + 0x88   image-slot discriminator: 2 = empty, 1 = set
record + 0x90   image sub-record
record + 0x180  cached GFX image (sub-record + 0xF0)
record + 0x4D8  dword given `(value & ~6) | 1` just before the commit
record + 0x4DC  flag byte set to 1 just before the commit
```

The transfer at `OART +0x22BCF4` is called with `rbp+8`, which is `record+0x88`:
the discriminator, followed by the image sub-record at `record+0x90`. The live
record prefix dump shows exactly that, `record +0x88 = 0x1` after transfer.

The transaction constructor is small:

```c
// OART +0x48870, called at +0x89CA48
Transaction(Transaction* self,          // stack storage, 0x510 bytes
            const PropertyRecord* src,  // rbp-0x80
            uint32_t flags,             // 0
            bool flag,                  // handler->byte_at_0x60
            uint32_t identifier);       // 0xA042008E
// self+0x00 vtable OART +0x9ED7E0, self+0x08 null singleton, self+0x10 = 0,
// self+0x14 = 1, record initialised in place at self+0x18,
// self+0x500 flags, self+0x504 identifier, self+0x508 flag
```

So the transaction is a thin stack value around the record, and the identifier
0xA042008E that later reaches the operation constructor is a literal in the
handler, not derived from the image.

Resolving the receiver is two dereferences. `OART +0x63EA0` is 0x16 bytes:

```c
receiver = *(void**)(token + 0x10);   // token = *(void**)(handler + 0x58)
```

The handler itself is the OART object behind PowerPoint's `FillFormat`; its
automation-facing wrapper is `OART +0x8A13E0`, which returns `0x800A01A8` on
rejection and otherwise calls `+0x89C860`. Before doing anything else, the
handler calls `receiver->vtable[0x130](receiver, &out)` and builds from the
result, which reads as a query of the Shape's current fill state.

That shape - query current state, mutate the image slot, construct a transaction,
commit - is the plausible outline of a native apply. It is an outline, not a
plan: the record is roughly 0x4E0 bytes built by 0x2BF bytes of handler code, and
only the four offsets above are mapped.

## Ownership and lifetime

The cached GFX image uses a 32-bit intrusive count at object+8, incremented
through vtable slot 0 and decremented through vtable slot +8 (GFX +0xA1A0 runs
`lock xadd` on +8 and destroys through slot +0x30 at 1).

Static chain, both directions:

* Property-record copy `OART +0x13360` copies the image sub-record at
  source+0x90 through `OART +0xA0870`, which AddRefs at `OART +0x9848C`.
* Property-record destructor `OART +0xAF80` destroys the image sub-record via
  `OART +0x3D870`, which Releases the cached pointer at `OART +0x3DBAB`,
  returning to **`OART +0x3DBB1`**. That is the exact return address recorded
  for the final Release during `Presentation.Close` in `resource_lifetime.md`,
  so the Close-time release and the operation-time release are the same code.
* Scalar deleting destructor `OART +0x66C80` destroys the embedded record at
  this+0x68 first, then releases this+0x50 and this+0x18, then frees the object
  through allocator slot +8 (observed target `ppcore.dll+0x2E1B70`).

Measured counts, identical in two independent runs:

| Sample point | Count |
|---|---:|
| Loaded image record | 1 |
| Transaction entry | 3 |
| Before construction (OART +0xEF3B) | 3 |
| After construction (OART +0x21BC06) | 4 |
| Before destruction (OART +0x21BC26) | 6 |
| After destruction (OART +0x66CDE) | 5 |
| UserPicture returns (OART +0x89CA71) | 5 |

**The operation owns a counted reference, it does not borrow one.** Construction
delta is exactly +1 and destruction delta is exactly -1. The apply step adds a
further +2 that outlives the operation; those are the document/undo references
that survive to `Presentation.Close`.

## Validation and remaining work

The harness compares Shape ID, name, type, bounds, rotation and Z order before
and after tracing. Both runs preserved those values and produced a picture fill.
This validates ordinary `UserPicture` under observation, not a synthesized
transaction or an optimized backend.

Answered by this experiment:

1. Which function constructs the operation — `OART +0x10244`, reached from the
   factory `OART +0xE7F0` at call site `+0xEF3B`.
2. What it receives — a property record plus flags/bool/identifier, all read out
   of the transaction; nothing else.
4. Whether the operation owns or borrows the cached image — it owns one counted
   reference for its own lifetime.
5. Part of lifetime — creation, apply and destruction are all inside the single
   `UserPicture` call; two further references survive it. Undo and Redo have not
   been sampled yet.

3. Where Shape identity lives — the receiver, which is constructed per Shape.
   The record prefix is Shape-independent apart from one unidentified pointer.

Still open:

6. Whether a cached image built from our own `IStream` can enter this flow. Both
   halves are now scoped rather than unknown: the receiver is reachable
   (`receiver_lookup.md`) and the record needs four slots set on top of Office's
   own constructors (`record_construction.md`). What remains is validating each
   private function's ABI before any of them is called.
7. Save/reopen of a synthesized fill — untested; only the ordinary path is
   validated.
8. Guarding. All anchors here are byte-validated per build, but no private call
   is enabled, so no compatibility profile has been widened.

No performance claim, production capability or compatibility profile changed.
File-free loading, donor-free assignment, independent resource lifetime and
repeated-decode avoidance remain unimplemented.

Raw evidence: `docs/evidence/fill_transaction_lifecycle.txt`,
`docs/evidence/fill_transaction_lifecycle_repeat.txt`,
`docs/evidence/fill_transaction_validated.txt`,
`docs/evidence/receiver_identity.txt` and
`docs/evidence/decoder_shape_validation.txt`. Saved package details:
`docs/fill_transaction_package.md`.
