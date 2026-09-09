# Resolving the per-Shape OART receiver

Observed on Office 16.0.14334.20848 x64, 2026-09-09. This document covers one
question: how the `Fill.UserPicture` handler obtains the OART receiver that the
fill transaction is applied to, and whether that receiver is reachable without a
debugger.

The answer is that it is reachable, through four pointer loads from the public
`Shape.Fill` object, each guarded by a vtable identity check. No private Office
function is called anywhere in this document.

## The chain

```text
Shape.Fill                 PPCORE FillFormat   vtable ppcore.dll+0x1464478
  + 0x08                   OART   FillFormat   vtable oart.dll+0xAF60B8
      + 0x58               control block       strong/weak counts, pointee
          + 0x10           OART   receiver     vtable oart.dll+0x9F6658
```

### Why each link is believed to be what it is

**PPCORE FillFormat -> OART FillFormat.** `PPCORE +0x8249C0` is slot 17 of the
public vtable and is a bare forwarder:

```c
void slot17(PublicFillFormat* self) {
    void* inner = self->at_0x08;
    inner->vtable[0x88](inner);     // 0x88 / 8 == slot 17
}
```

The inner vtable slot at +0x88 is exactly where the OART FillFormat vtable holds
`OART +0x8A13E0`, the automation entry that performs `UserPicture`. That entry
was already the observed top of the OART side of the live stack.

**OART FillFormat.** `OART +0xAF60B8` is an IDispatch vtable: slot 0 compares
IIDs against `OART +0xA066F0` and `+0xAF65F8`, slot 1 is
`++*(uint32_t*)(this+0x30)`, slot 2 the matching decrement that destroys through
the second base at `this+8`. So the reference count is a non-atomic uint32 at
+0x30, which is consistent with STA-only use.

Its factory is `OART +0x23B220`: allocate 0x68 bytes, install `+0xAF60B8` at +0
and `+0xAF47D0` at +8, AddRef a control block and store it at +0x58, then store
two flag bytes at +0x60 and +0x61. The byte at +0x60 is the same one the handler
later passes to the transaction constructor.

**Control block.** `OART +0x63EA0` is sixteen bytes long and reduces to
`return *(void**)(token + 0x10)`. The block is the ordinary MSO shape - strong
count at +0, weak count at +4, pointee at +0x10 - the same layout the operation
constructor builds inline at `OART +0x1037B`.

**Receiver.** Built by `OART +0x223950`, which takes exactly one argument beyond
`this`: the slide-level container, stored at +0x8. It also writes vtable
`+0x9F6658` at +0, the context sub-object vptr `+0x9F64D0` at +0x18, a flag byte
at +0x30, an empty small-buffer string at +0x38/+0x48, and at +0x20 a
**process-global allocation sequence number** taken with `lock xadd` from
`[OART +0xD40038]`. `OART +0x1C0C44` is the same constructor with a null
container.

## Corrections to earlier conclusions

Two statements in `fill_transaction.md` were under-determined and are corrected
here.

**`receiver+0x20` is not a Shape index.** It is an allocation counter written
once at receiver construction. It looked like an identifier only because each
Shape got its own receiver.

**"The receiver is per Shape" was not yet proven when it was first written.**
The original experiment filled three different Shapes, so a per-call object and a
per-Shape object were indistinguishable. `tools/prepare_receiver_identity.ps1`
now fills shape A twice before moving to B and C, which separates them. The
repeat is what makes every classification below sound.

## Measured classification

`experiments/exp_internal_blip/probe_receiver_identity.py` over the call order
A, A, B (same slide), C (other slide):

| Value | Classification |
|---|---|
| handler state (OART FillFormat) | stable per Shape |
| token (control block) | stable per Shape |
| receiver | stable per Shape |
| receiver +0x8 container | stable per slide |
| receiver +0x20 sequence | stable per Shape (assigned at construction) |
| cached GFX image | differs per call |
| property record +0x40 | differs per call |

`record+0x40` was the last candidate for Shape identity inside the record. It is
not: across two runs it was neither Shape-stable nor slide-stable, and the two
runs disagree in a way only heap-address reuse explains. Nothing in the property
record identifies the Shape.

`receiver+0x30` reads as a per-Shape value in the classifier, but the constructor
writes only a single flag byte there; the rest of that qword is padding that
happens to be uninitialised. Do not read meaning into it - the disassembly is the
evidence, not the classifier.

## Reachable without a debugger

`experiments/exp_internal_blip/receiver_inspect.cpp` performs the same walk from
inside the add-in and is exercised by `tools/test_receiver_lookup.ps1`. It
verifies both module versions, then checks three vtables in order, refusing to
dereference further the moment one does not match. It reads only; it takes no
reference and retains nothing.

Measured, agreeing with the debugger classification:

```text
A (first lookup)  receiver=0x229275c9790 container=0x22927678b40 sequence=2
A (second lookup) receiver=0x229275c9790 container=0x22927678b40 sequence=2
B (same slide)    receiver=0x229275caf30 container=0x22927678b40 sequence=4
C (other slide)   receiver=0x229275cb320 container=0x22927678c30 sequence=7
A (after an ordinary UserPicture) receiver=0x229275c9790  -- unchanged
```

The sequence numbers 2, 4, 7 are the same values the debugger run recorded for
the same Shape layout, from a completely independent code path.

Each lookup re-evaluates `shape.Fill`, so the stability is a property of the
Shape rather than of one COM wrapper instance. PowerPoint returns the same
`publicFill` pointer each time as well.

## The PPCORE vtable is validated structurally, not by address

The first version of this guard compared `*Shape.Fill` against one PPCORE RVA,
`+0x1464478`. That was too narrow, and a report of it rejecting a build turned
out to be the guard working correctly on the wrong object: the reported
`ppcore.dll+0x146EA20` is the **`Shape`** vtable, not `FillFormat`. Reading
`Shape+0x08` as an OART FillFormat would have been a wild pointer, so the
rejection was right - but the message did not say why.

Two things came out of investigating it.

**PPCORE has a family of these wrappers, not one class.** Scanning `.rdata` for
vtables whose slots are identity thunks found eleven, including `Shape.Fill`
(+0x1464478), `Shape.Line` (+0x1464338) and `Shape.TextFrame` (+0x14656A0). Each
delegates through `this+0x08` to its own OART counterpart. Pinning one RVA
therefore rejects legitimate siblings and would break on any build that moves the
table.

**The delegating shape is a stronger check than the address.** A wrapper's slot N
is a thunk of exactly this form:

```text
48 83 ec ??             sub  $imm8,%rsp          (optional)
48 8b 49 XX             mov  XX(%rcx),%rcx       inner = this->at_XX
48 8b 01                mov  (%rcx),%rax         its vtable
48 8b 80 YY YY YY YY    mov  YY(%rax),%rax       the slot  (or 48 8b 40 YY)
ff 15 ...               call *disp32(%rip)
```

`DescribeDelegatingWrapper` decodes those thunks, requires at least six of them
to map slot N onto inner vtable offset `N*8`, and requires every one to agree on
the same inner offset. So `innerOffset` is **read out of the code** rather than
assumed to be 8. The inner object then still has to present the exact OART
FillFormat vtable `+0xAF60B8` - which is what separates a Fill from the
identically shaped Line and TextFrame wrappers.

Measured, with the structural guard in place:

| Object | Result |
|---|---|
| `Shape.Fill` | accepted, wrapper +0x1464478, inner offset 0x8, 31 thunks |
| `ShapeRange.Fill` | accepted |
| `Shape` | rejected: vtable +0x146EA20 has no delegating thunks |
| `Shape.Line` | rejected: delegates to oart+0xAF4E38, not the FillFormat vtable |
| `Shape.TextFrame` | rejected: delegates to oart+0xAF5F48 |
| `Slide` | rejected: no delegating thunks |

Unknown layouts stay fail-closed; the difference is that the failure now names
the actual vtable, the inner offset it found, and what it expected.

`InspectFillReceiver` reports the environment alongside every result and attaches
it to every failure: expected build, the live `oart.dll`, `ppcore.dll` and
`gfx.dll` versions, the expected OART vtable RVAs, the PPCORE vtable actually
found, the derived inner offset and the thunk count. Build drift should be
readable straight off the failure message.

## Hazards

**A deleted Shape still resolves.** After `Shape.Delete()`, the walk still
returns a receiver and every vtable check still passes. The test records this
rather than asserting otherwise. Any future backend must therefore either
re-resolve the receiver on each apply - four loads, so this is cheap - or tie a
cached receiver to a Shape lifetime signal. Caching a receiver across a deletion
would be a use-after-free.

**STA only.** The OART FillFormat reference count at +0x30 is a non-atomic
`inc`/`dec`, so nothing in this chain may be touched off the owning thread.

**Everything here is borrowed.** No reference is taken anywhere in the walk. The
receiver is owned by the control block, which is owned by the OART FillFormat,
which is owned by PowerPoint.

## What this does and does not unblock

It answers where Shape identity lives: in the receiver, which is the OART side of
the Shape, reachable from the public `Shape.Fill`. It does not by itself make a
native apply possible. That still needs a valid property record, and the only
functions known to build one are inside the `UserPicture` handler. No capability
changed and no private call is enabled.

Evidence: `docs/evidence/receiver_identity.txt`,
`docs/evidence/receiver_lookup.txt`.
