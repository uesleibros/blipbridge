# Making the native apply faster

A cached `BB_ApplyTexture` cost about 0.19 ms. This is what that time turned out
to be made of, what could be taken off it, and what could not.

Everything below was measured on Office **16.0.14334.20848** x64, Release builds,
in process, never under a debugger. Absolute numbers move by a factor of two or
more with what else the machine is doing - a browser and a game running put the
same apply at 0.44 ms - so every comparison here is between legs measured in the
same run, and the ratios are what carry the claims.

## 1. Where the time goes

`experiments/exp_internal_blip/apply_profiler.cpp`, 5000 iterations, one
200x200 AutoShape, timing each step of the exact sequence `BB_ApplyTexture`
walks. Stage sum 0.18135 ms against a measured total of 0.18158, leaving
0.00024 ms unattributed, so the attribution is trustworthy.

| Stage | mean ms | share | whose code |
|---|---:|---:|---|
| virtualQuery (entry pointer check) | 0.00183 | 1.0% | ours |
| queryInterface | 0.00014 | 0.1% | Office |
| classify (Shape.Type, Shape.Connector) | 0.00777 | 4.3% | ours |
| fillFetch (Shape.Fill) | 0.00269 | 1.5% | Office |
| handleLookup | 0.00010 | 0.1% | ours |
| resolveTarget (the receiver walk) | 0.01234 | 6.8% | ours |
| resolveFunctions | 0.00004 | 0.0% | ours |
| record + subrecord + install + transfer + holder | 0.00494 | 2.7% | ours |
| transaction | 0.00246 | 1.4% | ours |
| **nativeApply** | **0.14826** | **81.6%** | **Office** |
| destructors | 0.00078 | 0.4% | ours |
| **total** | **0.18158** | | |

**The dominant cost is one call**: the receiver's own handler, which performs the
document edit. Four fifths of an apply is inside it and none of it is ours.

### The historical ~0.008 ms record-construction figure still holds

Re-verified under the v0.5.0 architecture rather than assumed: record, sub-record,
transfer, holder and destructors together are **2.96 us**; adding the cached-image
install and the transaction brings it to **9.2 us**. Between a third and a
twentieth of what the apply costs, depending on which boundary you draw. It is
not where the time is, so it was not optimised - see §7 for the pooling
experiment that was rejected on that basis.

## 2. What the 0.148 ms actually is

Two experiments, because "it is inside Office" is not an answer.

### It is not rendering

`apply_cost_factors.cpp` varies one thing at a time around the same apply, 2000
iterations each:

| Arrangement | mean ms | vs displayed |
|---|---:|---:|
| Shape on the displayed slide | 0.19011 | — |
| Shape on another slide | 0.18098 | -4.8% |
| Shape hidden | 0.19093 | +0.4% |
| Shape off the slide area | 0.18547 | -2.4% |
| Window minimised | 0.19337 | +1.7% |

Nothing moves it by more than 5%, and minimising the window makes it slightly
worse. So a batch that changes many Shapes before one redraw has no redraw cost
to amortise. That was worth knowing before building one.

### It is the edit itself, not bookkeeping around it

Receiver slot +0x78 (RVA 0x21DEC0) is a three-instruction thunk to slot +0x50
(RVA 0x21BB20), which stamps the transaction, tests a feature flag, and on the
off branch calls slot +0x60 (RVA 0x2290F0), then slot +0x58 (RVA 0x1B88B0), then
the produced object's own release. Driving those steps separately
(`transaction_split.cpp`, 2000 rounds) reproduces the combined route to within
0.008 ms and splits it:

| Step | mean ms | share |
|---|---:|---:|
| slot +0x60, compute the change | 0.00532 | 3.3% |
| slot +0x58, commit it | 0.14441 | 90.8% |

The tempting conclusion - that the commit is undo bookkeeping wrapped around a
5 microsecond edit, and could be skipped - is wrong, and the measurement says so.
A change computed and never committed **does nothing**: a clean Shape put through
it stays at `Fill.Type` 1 and its rendered PNG is byte-identical
(`tools/test_change_only.ps1`). The whole apply that way costs 0.0126 ms and
achieves nothing at all. The 144 microseconds are the edit.

## 3. What came off the individual apply

### Pinning the Office modules: the resolve went from 4.5 VirtualQueries to 2.5

`resolve_profiler.cpp` walks `ResolveFillTarget` step by step. Because
VirtualQuery's cost swings several-fold with machine state - 0.7 us in a trivial
process, 3 to 7 us inside PowerPoint, whose address space is not trivial - a
bare VirtualQuery is measured alongside as the normaliser.

Before:

| Step | mean us |
|---|---:|
| SynchroniseOfficeModules | 7.11 |
| readability checks and wrapper analysis | ~9.1 |
| whole resolve | 18.61 |
| (bare VirtualQuery) | 4.13 |

`SynchroniseOfficeModules` was three `GetModuleHandleW` calls, each taking the
loader lock and walking Office's module list by name, on every apply. They could
not simply be remembered: the handles were the validation cache's reload
detector. So they are **pinned** instead. A pinned module cannot be unloaded, so
its base cannot move and nothing else can appear at its address - a stronger
guarantee than the re-lookup it replaces, which could only notice a swap after it
had happened.

After:

| Step | mean us |
|---|---:|
| SynchroniseOfficeModules | 0.02 |
| whole resolve | 7.52 |
| (bare VirtualQuery) | 2.96 |

4.5 VirtualQuery-equivalents down to 2.5; 10.2% of an apply down to 6.8%.

What remains is almost entirely VirtualQuery itself - the readability checks that
stand between a wild pointer and a crash. That is the floor without giving one
up.

### The ceiling on this line of work

Everything outside Office's own call is 0.033 ms of a 0.182 ms apply. Removing
*all* of it - every guard, every classification - would reach 0.148 ms, and the
guards are not removable. So the individual apply is close to done, and the
remaining wins are about not making the call at all, or making one call do more.

## 4. Not making the call: `BB_ApplyTextureIfChanged`

Same acceptance, same gate, same document afterwards. The difference is that a
Shape already carrying the image is left untouched - no edit, no undo entry, no
invalidation.

5000 iterations per leg, 3 runs, medians of runs, on a quiet machine:

| Leg | mean ms | median | p95 | p99 | min | max |
|---|---:|---:|---:|---:|---:|---:|
| ApplyTexture repeated (baseline) | 0.18941 | 0.1767 | 0.2539 | 0.3860 | 0.1589 | 1.2383 |
| IfChanged, first apply | 0.18975 | 0.1777 | 0.2543 | 0.3889 | 0.1622 | 1.4150 |
| **IfChanged, repeated (skip)** | **0.02239** | **0.0210** | 0.0273 | 0.0508 | 0.0200 | 0.3979 |
| IfChanged, alternating A/B | 0.20266 | 0.1897 | 0.2698 | 0.4047 | 0.1721 | 1.3338 |
| ApplyTexture, alternating A/B | 0.20086 | 0.1885 | 0.2637 | 0.3936 | 0.1723 | 1.2932 |
| IfChanged, delete + recreate | 0.23092 | 0.2164 | 0.2902 | 0.5763 | 0.1973 | 1.3074 |

| Pass over 32 Shapes sharing one image | total ms | per Shape |
|---|---:|---:|
| 32 x ApplyTexture | 6.673 | 0.2085 |
| 32 x IfChanged, all skipping | 0.740 | 0.0231 |

- **8.5x**, saving 0.167 ms per redundant call.
- The cost when it never gets to skip is **4.9 us** - the Shape key.
- A first apply is indistinguishable from the baseline.

Repeating the same run on a loaded machine put the baseline at 0.545 ms and the
skip at 0.132 ms - every leg inflated by about the same factor, the ratio 9.5x.
That is why the ratios are quoted rather than the absolutes.

### What the skip is allowed to know

The remembered image is a **process-unique image id**, never the caller's handle:
a handle can be released while another owner keeps the same picture alive, so
handle equality answers a different question. Nothing is cached that could
outlive what it names - no receiver, no control block, no `FillFormat`, no Shape
pointer, only values - and `Fill.Type` is re-read before any skip is granted.

The record is **shared with `BB_ApplyPicture`'s cache** rather than duplicated
beside it. Two maps over the same Shapes could disagree - one saying a Shape
carries image X just after the other put Y on it - and the next skip would honour
the stale one and show a wrong picture. Every path that writes a fill now writes
the record; the ordinary apply only pays for the Shape key once something has
actually asked for a skip.

That also fixed a latent case in the shipped code: `BB_ApplyTexture` changed a
fill without telling the picture cache, so a following `BB_ApplyPicture` could
skip on the strength of an image that was no longer there.

### The existing batch API is not a batch

`BB_ApplyTextureBatch` saves one ABI entry, one thread check and one argument
validation instead of N. In process that is nothing, and the benchmark says so:

| Shapes | individual | batched | per Shape individual | per Shape batched | speedup |
|---:|---:|---:|---:|---:|---:|
| 10 | 1.752 | 1.848 | 0.1752 | 0.1847 | 0.95x |
| 50 | 8.667 | 8.638 | 0.1733 | 0.1728 | 1.00x |
| 100 | 17.722 | 17.629 | 0.1772 | 0.1763 | 1.01x |
| 200 | 35.135 | 35.070 | 0.1757 | 0.1753 | 1.00x |

It exists to cut language-boundary crossings, and a VBA caller does save the
`Declare` transition per call - but it does not make Office's work any smaller,
because it still makes N separate edits. That is the gap the next section
closes.

## 5. Making one call do more: the ShapeRange apply

A transaction carries no target - the receiver *is* the target - so a transaction
holding several Shapes' picture fills cannot exist. What can exist is a receiver
that stands for several Shapes, and `ShapeRange.Fill` is one: same PPCORE
delegating wrapper (`ppcore.dll+0x1464478`, inner offset 0x8, 31 thunks), same
OART FillFormat vtable (`oart.dll+0xAF60B8`), same validated receiver vtable
(`oart.dll+0x9F6658`), same container. The existing structural walk accepts it
unchanged and one apply fills every member.

### Measured

All three legs through the public C ABI, in process, same Shapes, same image, 3
runs per size with the median reported. `tools/run_range_benchmark.ps1`.

**Reference run, quiet machine** - the figures quoted elsewhere in this repo:

| Shapes | one at a time | `BB_ApplyTextureRange` | per Shape | speed-up |
|---:|---:|---:|---:|---:|
| 1 | 0.229 | 0.232 | 0.232 | 1.0x |
| 2 | 0.449 | 0.299 | 0.149 | 1.5x |
| 8 | 1.771 | 0.632 | 0.079 | 2.8x |
| **32** | **6.650** | **1.876** | **0.059** | **3.5x** |
| 100 | 20.935 | 5.549 | 0.055 | 3.8x |

Those are measurements from one machine in one session, not constants.

**Full sweep, busy machine** - a browser, a chat client and a remote-desktop
service running, which roughly triples every absolute figure:

| Shapes | one at a time | `BB_ApplyTextureBatch` | `BB_ApplyTextureRange` | per Shape | vs loop | vs batch |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 0.236 | 0.237 | 0.241 | 0.241 | 1.0x | 1.0x |
| 2 | 0.423 | 0.426 | 0.288 | 0.144 | 1.5x | 1.5x |
| 8 | 4.926 | 5.044 | 2.031 | 0.254 | 2.4x | 2.5x |
| 16 | 10.906 | 11.225 | 3.770 | 0.236 | 2.9x | 3.0x |
| 32 | 18.750 | 19.114 | 6.421 | 0.201 | 2.9x | 3.0x |
| 64 | 41.264 | 44.845 | 13.429 | 0.210 | 3.1x | 3.3x |
| 100 | 51.235 | 50.759 | 15.426 | 0.154 | 3.3x | 3.3x |

Every Shape filled in every row. The absolute numbers are three to four times the
quiet ones - the per-Shape leg reads 0.5-0.7 ms against a 0.19 ms baseline - and
the **ratios barely move**, which is the point of measuring all three legs in the
same run. Quote the ratios; re-measure the absolutes.

`BB_ApplyTextureBatch` tracks the one-at-a-time leg at every size, because it
saves ABI crossings rather than Office work.

The comparison legs have to run in process. Driven from PowerShell they read
about 5.9 ms per Shape - a cross-process Automation round trip, not Office - and
make the range look 25 to 155 times faster for reasons that are nothing to do
with the work being compared.

### The gate is not optional

The fill reaches every member, so a Connector in the range would reach the
private backend, and a native picture fill on a Connector terminates PowerPoint.
Every member is classified before anything internal is touched, and one
ineligible member refuses the whole range by name. `tools/test_range_apply.ps1`
proves the refusal happens *before* the private apply using the apply-entry
counter, for a Connector, a Line and a Table, and proves the other members are
left as they were rather than half-filled.

### What a range may hold

One range mixing AutoShape, TextBox, WordArt, Freeform and Callout fills all
five. The two worth naming are WordArt, which reports `msoAutoShape` - the same
type a Connector reports - and the Freeform, which reports a ShapeRange as its
parent; both have caught something before.

A group in the range is **filled, not refused**: `msoGroup` has a validated
native path. It is never skipped afterwards, because filling a group changes what
its children render and filling a child changes what the group shows, so neither
is ever keyed.

A ShapeRange belongs to one slide's `Shapes` collection, so there is no
cross-slide range to hand this path. Filling several slides means one call per
slide, and the ~0.12 ms fixed cost is paid once per slide.

### Undo is one entry, like Office's own

Counted rather than guessed at: a marker Shape goes on the undo stack first, then
Undo is repeated until the marker disappears, which gives the number of entries
an operation added.

| Operation | 4 members | 8 | 16 |
|---|---:|---:|---:|
| nothing | 0 | 0 | 0 |
| N x `BB_ApplyTexture` | 4 | 8 | 16 |
| **`BB_ApplyTextureRange`** | **1** | **1** | **1** |
| Office's own `ShapeRange.Fill.UserPicture` | 1 | 1 | 1 |

One range apply is one undo entry, one Undo reverts the whole fill, and one Redo
restores it - the same as the public Office operation, at every size tested and
for a range of five different Shape classes. The path that leaves an entry per
Shape is the per-Shape one, which is what it has always done.

This measurement was wrong twice before it was right, both times for the same
reason: it was not counting entries, it was watching fills. Pressing Undo five
times against a six-entry stack looked like "no undo at all", and a later count
of N + 2 was taken while the *benchmark* harness was under test - which also
applied to every member individually, and left those entries behind. Counting
against the public entry point, with a marker, gives the table above.

Undo and Redo change fills without telling BlipBridge, like any change made
outside it, so the per-Shape record can be stale afterwards and
`BB_InvalidateShape` is the remedy - the same rule as for any external change.

### It is a public API

`BB_ApplyTextureRange(shapeRange, texture, &applied)`, ABI 4, with
`BlipBridge.ApplyTextureRange(shapes As PowerPoint.ShapeRange, texture)` in the
VBA wrapper. `docs/c_abi.md` has the contract; the section above has the reason
the argument is a ShapeRange and not a list of Shapes.

## 6. What is production-safe

| Change | Status |
|---|---|
| `BB_ApplyTextureRange` | Ready. New export, ABI 4, 46 assertions in `tools/test_range_apply.ps1` |
| `BB_ApplyTextureIfChanged` | Ready. New export, 33 assertions in `tools/test_apply_if_changed.ps1` |
| One shared per-Shape record | Ready, and fixes a latent wrong-picture case in the shipped code |
| Pinned Office modules | Ready. Strictly stronger than what it replaces |
| Freeform keying | Ready. `ConvertToShape` reports a ShapeRange as its Parent; those Shapes could never be keyed, in either cache |
| `ApplyRoute::Split` / `ChangeOnly` | Research only. Never reachable from the C ABI |

## 7. Rejected, and why

- **A per-call region memo for the readability checks.** The four objects in the
  receiver chain occupy three distinct memory regions, so it would remove one
  VirtualQuery of four - not worth the reasoning it would cost a reader.
- **Prepared or pooled records.** Record construction is 2.96 us and the whole
  record-and-transaction group is 9.2 us, against 148 us for Office's call.
  Pooling would trade a measurable ownership hazard for an unmeasurable saving.
- **Skipping the commit (`ChangeOnly`).** 13x faster and does nothing: the Shape
  stays unfilled and renders identically. Kept only so the negative result stays
  reproducible.
- **Batching before one redraw.** There is no redraw cost to amortise; see §2.
- **A transaction holding several Shapes' fills.** The transaction carries no
  target, so this cannot exist. The multiplicity has to live in the receiver.

## 8. Recommendation

Everything in §6 is ready. The question is only whether it goes out in one
release or two.

**Recommended: one release, v0.6.0, ABI 4, with the changelog separating the two
halves** - which it does. Two releases would mean running the whole validation
matrix twice for the same code, and the correctness fixes are not independently
urgent: the wrong-picture case needs a caller that mixes `BB_ApplyTexture` with
`BB_ApplyPicture` on the same Shape, and the Freeform one costs a redundant apply
rather than a wrong result. Neither is a reason to hold the release, and neither
is a reason to ship twice.

If a v0.5.1 is wanted anyway - because someone is hitting the mixed-API case
today, or because separating a fix release from a feature release is worth
something on its own - the two fixes are independent of the range work and can be
cherry-picked: the shared per-Shape record and the Freeform key. Both land in
files the range API does not touch.

### Known, and deliberately not fixed here

The undo *entry* matches Office, but BlipBridge cannot see a fill that Undo or
Redo changes, so the per-Shape skip record can be stale after one.
`BB_InvalidateShape` is the documented remedy, and it is the same rule that
already applies to a fill changed by another add-in, a paste or a theme change.

### Future research

**Investigate PPCORE ShapeRange undo coalescing.** Not a blocker any more - the
range path already produces one entry - but the mechanism is still unknown, and
knowing it would settle whether the match is guaranteed or incidental.

Where a first pass stopped, so a second does not repeat it: the commit
(`oart.dll+0x1B88B0`) gates its journalling on the receiver's `vtable[+0x48]`
(`oart.dll+0x1B87D0`), which opens with a *global* test - `cmpq $0,
[oart+0xD4AEA8]` - rather than anything per receiver, and the branch taken when
that fails leads to an assertion helper rather than an alternative recording. So
the behaviour is not a flag on this path. Following it means following PPCORE.
