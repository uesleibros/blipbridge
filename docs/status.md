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

Next experiment: trace that consumer into document-resource/fill binding and verify balanced lifetime dynamically before attempting a native memory-stream factory. Determine whether the GFX object is only a rendering cache. A direct decoder-stream substitution alone would leave the earlier temporary PNG write intact. Production cache and VBA interpreter overhead remain open.
