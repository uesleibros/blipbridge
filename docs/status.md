# Current status

Current backend: **PickupApplyFallback**, experimental, whole-style transfer from an existing normal donor shape with a preloaded picture fill.

A reusable native texture API now exists alongside it - LoadTexture,
ApplyTexture, ReleaseTexture, ClearTextures - decoding bytes once into a GFX
cached image and applying it with no file, no donor Shape and no UserPicture.

Current position, each item measured:

  native cached apply                        proven
  file-free decode, no temporary image       proven
  reuse across Shapes, slides, presentations proven
  save / reopen                              proven
  performance win                            proven
  Undo / Redo                                proven
  reference ownership                        fully attributed

Benchmark, in process over 1000 iterations: ApplyTexture 0.1860 ms mean /
0.1756 median / 0.2260 p95 / 0.3361 p99, against Fill.UserPicture's 0.6588 /
0.6189 / 0.8348 / 1.6002. Mean speed-up 3.54x, and the p99 gap is larger still.
Alternating two textures costs the same as repeating one. LoadTexture 0.0231 ms,
ReleaseTexture 0.0001 ms.

Capabilities are now computed at runtime rather than hard-coded, because the
backend is pinned to one Office build and must fail closed elsewhere. Inside
PowerPoint on the validated build they read MemoryImageToFill=True,
CachedTextureApply=True, InternalBackend=True; outside, all three read False.
FillOnly stays False deliberately - see capabilities.md. See native_texture.md
for the evidence behind each claim.

2026-09-09 code-quality checkpoint: COM activation, dispatch, donor operations,
and research routing are separated and documented. Named DISPIDs/HRESULTs and
exception-string RAII are in place. Release/Debug builds and native COM contract
tests pass; PowerPoint fallback, memory and smoke regressions pass. New coverage
includes wrong-thread rejection, server unload/locks, invalid shape types,
double release, missing arguments and unchanged capabilities. Further cleanup of
the older image validator and IAT instrumentation remains incremental work.

Milestone 1: GCC/UCRT64 x64 Release executable and Automation DLL built. Office exact build detected. Reproducible environment probe records loaded modules.

Milestone 2: full visible-document in-process benchmark completed, including 100/1,000/10,000 distributions, surrounding operations and 130 pooled freeforms. Separate hidden-document dual-COM timings completed. External baseline is partial due to a rejected COM call.

Milestone 3: actual UserPicture file/buffer boundary traced through PPCORE, OART and MSO20, with RVAs and focused disassembly. Complete decoder/resource-binding ABI remains unresolved. Exact Office PDB requests returned HTTP 404.

Milestone 4: **Not achieved under the strict no-temporary-image definition.** Byte arrays reach existing freeform fills through a version/signature-guarded Win32 adapter, but expanded tracing proved Office writes and rereads a PNG under INetCache/Content.MSO. Earlier "RAM-only" wording described source-byte delivery too broadly. PNG sizes 32/64/128/256 and JPEG worked; identity, editable nodes, pixel equality and save/reopen checks remain valid. Normal LoadTexture/SetImageBytes remain E_NOTIMPL.

What works: retaining donor shapes through safe COM references; fallback ApplyTexture; geometry/identity preserved in the first freeform test; editable nodes after transfer; DLL runs in POWERPNT through research add-in.

What does not work: production direct-byte API, decoded BLIP ownership, fill-only transfer, internal resource cache. No undocumented internal function pointers are called.

Known errors: external COM benchmark RPC_E_CALL_REJECTED during shape churn. No confirmed native crashes to date.

Current measurements: hidden dual-COM UserPicture ~1.09 ms; pre-picked Apply ~0.87 ms; PickUp+Apply ~1.25 ms. Visible 130-fill/node/visibility batch: UserPicture ~1.28 s/frame, PickUp+Apply ~2.15 s/frame, with fixed run order and accumulated document history. These are not refresh FPS. Full distributions are in docs/benchmarks.md.

Reliability: 100 alternating experimental byte assignments completed (~2.44 ms before WIC validation was added). 100,000 alternating cached-donor applications passed (~2.44 ms mean across the long run). Range batching was slower (~1.75 s per 130-shape range versus ~0.51 s for individual fills in the same history-heavy run). Private bytes peaked around 3.04 GB and returned to 225 MB after document close. This is significant document/undo state, not a proven native leak. Stale/released/deleted donors and invalid images are covered by COM smoke tests.

Latest continuation: Release/Debug builds and memory/COM regression tests passed. Stress deck reopened with 133 normal picture-filled shapes, zero Picture shapes and two media resources. Live decoder object inspection plus its QueryInterface GUID comparison confirmed a standard IStream beneath GFX +0x7880. GFX exports stream-based ICachedImage::Create (+0x7680) and IImage::Create (+0x194090); neither is called by BlipBridge.

Factory entry/return observation now captured distinct cached-image and image outputs (GFX vtables +0x409DC0/+0x4055C8), with private intrusive reference counts established statically. The existing factory returns to OART +0x8F554, which passes its result to +0x8F94C and releases a local reference.

2026-09-09 native progress: OART +0x8F94C stores the cached object in record +0xF0.
Live counts show retention followed by release of the caller's local reference.
Further AddRefs occur in OART record copying (+0x9848C), including the UserPicture
transaction after image loading. Phase-tagged tracing proves the cached object
survives UserPicture and SaveAs; its last observed release occurs during Close.
The ordinary traced deck saved/reopened with one normal image-filled AutoShape.
See resource_lifetime.md for evidence and remaining ownership limits.

Latest transaction trace: the cached pointer propagates from loaded record
+0xF0 to wrapper +0xF8, transaction +0x198 and operation +0x1E8. These are one
layout: the property record embeds an image sub-record at +0x90 whose cached
pointer sits at +0xF0. The receiver dispatch resolves through OART
+0x21DEC0/+0x21BB20, which is a build/apply/delete driver. See fill_transaction.md.

2026-09-09 operation lifecycle resolved. The operation is built by factory OART
+0xE7F0: 0x570 bytes from the Office allocator, then constructor OART +0x10244
with (property record at transaction+0x18, flags transaction+0x500, bool
transaction+0x508, identifier transaction+0x504). Observed identifier 0xA042008E,
flags and bool zero, resulting vtable +0x9E4BD0. The constructor receives no
Shape and no document pointer; identity is carried by the receiver
(vtable +0x9F6658, context sub-object at receiver+0x18), which OART +0x63EA0
resolves from handler state +0x58. The operation is applied by OART +0x1B88B0 to
that receiver and then destroyed by OART +0x66C80 within the same UserPicture
call, freeing through ppcore.dll+0x2E1B70.

Ownership resolved: the cached GFX image count (32-bit at object+8) measured
1 -> 3 -> 3 -> 4 -> 6 -> 5 -> 5 across loaded record, transaction entry, before
and after construction, before and after destruction, and UserPicture return.
Construction is exactly +1 and destruction exactly -1, so the operation owns a
counted reference rather than borrowing one; the two references that survive the
call are taken during apply and live until Presentation.Close. Both runs of
tools/run_office_probe.ps1 produced identical values and preserved Shape
identity, geometry, Z order and picture fill.

Receiver scope measured. Filling three AutoShapes - two on one slide, one on
another - with identical bytes produced three distinct receivers, handler states
and tokens, so the receiver is constructed per Shape. receiver+0x8 is a PPCORE
object (vtable ppcore.dll+0x1396DB8) shared by Shapes on the same slide and
different across slides, consistent with a slide-level container. The property
record prefix below the image sub-record was identical for all three Shapes apart
from one unidentified non-polymorphic pointer at record+0x40, which also tracks
the per-call image resource and so cannot yet be read as Shape identity.

Consequence: a native ApplyTexture cannot be reached by constructing a property
record alone, because neither the record nor the operation names the target
Shape. The open problem is obtaining the per-Shape OART receiver from a
PowerPoint Shape without going through the UserPicture handler; nothing observed
so far shows that is reachable.

Static follow-up on the same day: the receiver is two dereferences from the fill
handler (OART +0x63EA0 is `receiver = *(token + 0x10)`, token = handler+0x58),
and the transaction is a thin 0x510-byte stack value built by OART +0x48870 from
(property record, flags 0, handler byte +0x60, identifier 0xA042008E). The
property record is the ~0x4E0-byte object the whole handler builds in its own
frame; the image-slot discriminator at record+0x88 and the image sub-record at
record+0x90 are inside it, which the live prefix dump corroborates. The handler
starts by calling receiver vtable+0x130 to query current state. Neither the
transaction nor the receiver lookup is the hard part; the unmapped property
record and reaching the handler from a Shape are.

Receiver resolution solved. Shape.Fill is a PPCORE object whose vtable slot 17
(PPCORE +0x8249C0) forwards through this+0x08 to the OART FillFormat, so the
chain is Shape.Fill -> +0x08 -> OART FillFormat -> +0x58 -> control block ->
+0x10 -> receiver, with vtables ppcore+0x1464478, oart+0xAF60B8 and
oart+0x9F6658 as guards. The read-only InspectFillReceiver research method walks
it from inside the add-in and reproduces the debugger classification, including
the same allocation sequence numbers. Hazard recorded: a Shape deleted through
public COM still resolves through the whole chain, so a cached receiver would be
a use-after-free; re-resolving costs four loads. Correction: receiver+0x20 is a
process-global allocation counter, not a Shape index. See receiver_lookup.md.

Record construction mapped. The handler builds the whole ~0x4E0-byte record with
Office's own constructors (+0x14110, +0x14F7B0, +0x14580) and then sets only four
things: the slot at +0x00 with payload 3, the image slot at +0x88/+0x90 via the
transfer, the counted 16-byte slot at +0x2A0, and the tagged flag at
+0x4D8/+0x4DC. +0x14F7B0 enumerates the record's slot table. So a synthetic
record need not be fabricated byte by byte. The stream-based cached-image creator
is a real GFX export (ordinal 236, RVA 0x7680) whose mangled name and prologue
give the full ABI, including that RCX is the hidden sret pointer and the returned
pointer carries one reference. See record_construction.md.

Measured for the performance goal: the cached GFX image differed on all four
observed UserPicture calls, including two on the same Shape with the same file.
Office builds a new cached image per call and reuses nothing.

First native call made, and it is the load half only. LoadCachedImageExperiment
calls the exported GFX ICachedImage::Create by name and releases what it creates.
It touches no Shape, slide, presentation or property record. PNG and JPEG bytes
decode from memory with no file on disk and no UserPicture. Every observed field
matched the static prediction: cached count 1, image count 2, vtables GFX
+0x409DC0 and +0x4055C8 - the same numbers resource_lifetime.md recorded by
watching Office, now reproduced from our own call. Invalid bytes are rejected
without a crash and ordinary UserPicture still works afterwards. Nothing is
retained, so this is not yet a texture handle and there is no performance claim.

Native apply works. NativeApplyExperiment applies image bytes to an existing
normal Shape with no UserPicture on the target, no donor Shape, no PickUp/Apply
and no source file: the bytes reach Office through a memory IStream, become a
cached image through the GFX export, and are committed through Office's own
record, transaction and receiver. Measured on an AutoShape and a Freeform:
Fill.Type 1 -> 6, Shape id/name/geometry/rotation/Z order unchanged, still a
normal AutoShape with zero Picture shapes, Freeform kept three editable nodes,
SaveAs and reopen both preserved the fill, and sampled pixels are identical both
to an ordinary UserPicture fill and after reopen.

Getting there needed one correction found by measurement: with the counted slot
at record+0x2A0 left unset the apply produced Fill.Type 4 (textured) instead of 6
(picture). That slot is exactly what separates the two - the FillFormat vtable's
next entry OART +0x8A1530 (UserTextured) calls the same handler +0x89C860 and
differs only in building this argument with +0x239950 instead of +0x158C40.

Thirteen private OART entry points are now called, each byte-verified at its RVA
before use, with sizes, ownership and destructor pairing recorded in oart_abi.md.
Record 0x4E8, image sub-record 0x200, transaction 0x510. The receiver is
re-resolved before every apply and never cached.

Cached-image reuse works. NativeApplyReuseExperiment decodes once and applies
that single cached image to five Shapes: two plain AutoShapes and a different
AutoShape type on slide 1, one on slide 2, and a Freeform. The report shows
creations=1 with one cached-image address and five distinct receivers. Every
Shape kept its identity and got a picture fill, the Freeform kept three editable
nodes, and reopening the saved deck found six Shapes with picture fills and zero
Picture shapes. Applying repeatedly to one Shape also preserved identity and the
fill. So the same GFX cached image can serve many Shapes without Office building
a new one, which is the reuse the whole design depends on.

The reference delta is characterised, and retention is measured. Across the
receiver call the count gains two or three references depending on document
state - three consecutive identical applies measured +3, +2, +3 - and it does not
grow with repeated applies. Two hundred repeated native applies to one Shape
moved the host's private bytes from 98.0 MB to 98.0 MB, and 98.1 MB after closing
the presentation, so nothing is being stranded. The individual references remain
unattributed; that attribution is a productionization prerequisite rather than a
blocker for research.

Stability: 200 repeated applies preserved Shape identity, geometry and the
picture fill; ordinary UserPicture worked before and after; a fresh presentation
worked afterwards. Undo could not be exercised - CommandBars.ExecuteMso('Undo')
returns E_FAIL in this automation harness after an ordinary UserPicture too, so
the result is inconclusive for both paths rather than a finding about the native
apply. The host survived every attempt. Whether a native apply registers an undo
entry is therefore still unknown; the real handler's wrapper OART +0x8A13E0 sets
up an action scope that the experiment deliberately skips.

Capabilities stay false. The experiments retain nothing across calls, so there is
still no texture handle and no performance claim: reuse has been shown to work,
but nothing has been timed against Fill.UserPicture yet. Internal cache-file elimination and donor-free
fill binding remain unresolved; no private call is enabled, and no capability
flipped.
