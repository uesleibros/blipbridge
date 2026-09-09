# Current status

Current backend: **PickupApplyFallback**, experimental, whole-style transfer from an existing normal donor shape with a preloaded picture fill.

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
tools/run_fill_transaction.ps1 produced identical values and preserved Shape
identity, geometry, Z order and picture fill.

Next experiment: identify which receiver field denotes the target Shape, and map
the property record below +0x90. Both are prerequisites for supplying a
stream-created resource. Internal cache-file elimination and donor-free fill
binding remain unresolved; no private call is enabled, and no capability flipped.
