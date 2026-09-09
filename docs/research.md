# Research journal

## 2026-09-09 - record, transaction and receiver lookup are all small

Follow-up static pass on OART +0x89C860, with no new live run.

The receiver lookup is trivial: OART +0x63EA0 is 0x16 bytes and reduces to
`receiver = *(void**)(*(void**)(handler + 0x58) + 0x10)`. The transaction is a
0x510-byte stack value built by OART +0x48870 from (property record, flags,
bool, identifier); the identifier 0xA042008E that reaches the operation
constructor is a literal at the call site, not derived from the image.

The pieces observed separately turn out to be one record built in the handler's
own frame at rbp-0x80: the transfer destination at rbp+8 is record+0x88, the
image sub-record is record+0x90 and the cached image record+0x180. The live
prefix dump agrees - record+0x88 reads 1 after the transfer, matching the
discriminator write. The handler's first act is `receiver->vtable[0x130]`, which
reads as querying the Shape's current fill state.

So the outline of a native apply would be query, mutate the image slot, build a
transaction, commit - but that is an outline, not a plan. The record is roughly
0x4E0 bytes produced by 0x2BF bytes of handler code and only four of its offsets
are mapped, and reaching the handler from a PowerPoint Shape is still unsolved.
Recorded so the next session does not re-derive it. No code changed.

## 2026-09-09 - the fill receiver is per Shape

Hypothesis: since neither the operation nor its property record carries a Shape
reference, the target must be denoted by the receiver. Its scope was unknown.

Method: `tools/prepare_receiver_identity.ps1` fills three ordinary AutoShapes -
two on slide 1, one on slide 2 - with the same PNG through the public
Fill.UserPicture, with no IAT instrumentation installed.
`probe_receiver_identity.py` records the handler state, the token at handler+0x58,
the receiver, a bounded receiver dump and the property-record prefix below the
image sub-record, then reports only what differed. Both scripts run through the
generalized `tools/run_office_probe.ps1`.

Result: the handler state, token and receiver are all distinct for every Shape,
so the receiver is per Shape rather than per slide or per document. receiver+0x8
holds a PPCORE object with vtable ppcore.dll+0x1396DB8 that is shared by the two
Shapes on one slide and different for the Shape on the other, which reads as a
slide-level container - inferred from the sharing pattern and a stable vtable,
not from a symbol. receiver+0x20 is a small per-Shape integer that is not
Shape.Id and not stable between runs. receiver+0x28..+0x48 is a small-buffer
string, with +0x38 pointing at +0x48.

The property-record prefix was byte-identical across all three Shapes except
record+0x40, a heap pointer whose first qword is zero. Each call also produced
its own cached image, so that difference cannot yet be attributed to Shape
identity; it stays unidentified rather than being assigned a role.

Consequence for the intended backend: constructing a property record is not
sufficient for ApplyTexture, because the record does not name a Shape. The
blocking problem is obtaining the per-Shape receiver outside the UserPicture
handler, and nothing observed shows that it is reachable.

Limits: two runs, one Office build, three Shapes, one image. Shape identity was
preserved and every Shape ended with a picture fill. No internal function was
called, no capability changed, no benchmark claim made.

## 2026-09-09 - operation construction, inputs and cached-image ownership

Hypothesis: the operation object handed to application slot +0x58 is built by a
locatable factory whose inputs and resource ownership can be established without
calling anything.

Method: static PE work first. `.pdata` lookups resolved every observed return
address to a real function entry, so the construction chain could be read from
disassembly before any breakpoint was placed. The probe was then rewritten
around those anchors, adding intrusive reference-count sampling at six points,
and driven twice through the new `tools/run_office_probe.ps1`.

Result. The construction chain is OART +0x2290F0 -> +0x21BC50 -> +0x2F1E00 ->
+0xF740 -> +0xE7F0 -> +0x10244, confirmed by the live backtrace in both runs.
The factory allocates 0x570 bytes from the Office allocator and constructs with
(property record at transaction+0x18, flags transaction+0x500, bool
transaction+0x508, identifier transaction+0x504). Observed identifier 0xA042008E;
flags and bool zero. The constructor copies the property record into
operation+0x68, so the cached image reappears at operation+0x1E8, which the probe
verified equal to the tracked GFX pointer.

The constructor receives no Shape or document argument. Identity travels with the
receiver (vtable +0x9F6658), whose context sub-object is at receiver+0x18 and
which OART +0x1B88B0 passes to +0x1B8F50 together with receiver+0x10. The driver
at +0x21BB20 is simply build, apply, destroy: the operation is freed inside the
same UserPicture call through +0x66C80 and ppcore.dll+0x2E1B70.

Ownership is now measured rather than inferred. Counts on the cached image were
1, 3, 3, 4, 6, 5, 5 at loaded record, transaction entry, before and after
construction, before and after destruction, and UserPicture return - identical in
both runs. Construction is +1 and destruction -1, so the operation owns a counted
reference. This also corrects an incomplete note in resource_lifetime.md: the
Close-time release at OART +0x3DBB1 is not a close-specific path, it is the image
sub-record destructor OART +0x3D870, whose matching AddRef is the record copy
OART +0x13360 -> +0xA0870 -> +0x9848C.

Limits: which receiver field denotes the Shape is still unknown, and the property
record below +0x90 is unmapped. Both must be resolved before a stream-created
resource could enter this flow. No internal function was called, no capability
changed, no benchmark claim made.

Validation: both traced runs preserved AutoShape ID, name, type, bounds, rotation
and Z order and produced a valid picture fill; all breakpoints were removed and
the debugger detached. Release and Debug builds, CTest, COM smoke and
fallback-contract tests passed afterwards.

## 2026-09-09 - loaded record reaches the state transaction

Hypothesis: the GFX cached reference is carried by the transaction used after
UserPicture image loading. Added an isolated read-only GDB probe with named,
version-specific signatures and bounded record observations. Captured source
at OART +0x22BCF4, transaction call at +0x89CA6B and its return.

Result: exact pointer equality at source+0xF0, destination-wrapper+0xF8 and
transaction+0x198. A 0x100-byte initial observation was too small; constructor
disassembly justified the later 0x510-byte transaction bound. Receiver vtable
+0x9F6658 slot +0x78 resolves to +0x21DEC0, which forwards to slot +0x50 at
+0x21BB20. Repeated live observations resolve preparation slot +0x60 to +0x2290F0
and application slot +0x58 to +0x1B88B0. Application receives a separate operation
object (vtable +0x9E4BD0), not the GFX image itself. No internal function called
by the probe. Exact command semantics and captured document/Shape ownership remain
unproven. Full observations, static corroboration and next experiment are in
fill_transaction.md.

Validation: repeated trace completed and all temporary breakpoints detached.
AutoShape ID/name/type/bounds/rotation/Z order stayed equal; picture fill remained
valid. Saved package has one normal picture-filled Shape, zero Picture shapes
and one PNG. Native COM tests passed in Release/Debug and fallback-contract
PowerPoint tests passed. Native implementation unchanged; no new benchmark claims.

## 2026-09-09 - OART consumer and resource retention

Hypothesis: OART +0x8F94C retains the GFX cached object in state used beyond image
loading. Added named/signature-checked consumer, AddRef and Release observations
to the existing GDB probe. No inferior function calls or memory writes beyond
temporary debugger breakpoints. Stop before any tracked count-1 release to avoid
observing freed storage.

Results: record +0xF0 changes from null to the exact cached pointer. Cached count
is 1 at factory return, reaches 2 before the caller's local release, and remains
1 afterward. Subsequent copy-path AddRefs return via OART +0x9848C; static code
copies record +0xF0 and increments that interface. Dynamic stacks connect these
copies to UserPicture after its loader (+0x89CA06/+0x89CA4D).

An initial lifetime run lacked phase markers, so close-time destruction could
not be attributed confidently. Repeated with explicit UserPicture/SaveAs/Close
markers: final cached Release count-before 1 occurs during Presentation.Close,
via OART +0x3DBB1. Thus the resource is not merely discarded when loading returns.
Which references belong to undo versus active document state is still unresolved.
The separate image's final destruction was not followed after this stopping point.

The traced ordinary UserPicture presentation saved/reopened as one normal
picture-filled AutoShape, one media PNG and no Picture shapes. No optimized
backend, direct GFX call or speedup is claimed. Findings/RVAs and next binding
experiment are in resource_lifetime.md; raw phase and retention stacks archived.

Post-experiment validation: Release/Debug builds and native COM contract tests
passed, as did the public fallback regression suite. Native backend code was not
changed in this investigation. The debugger detached and presentation close
completed after each successful observation.

## 2026-09-09 00:08 onward - incremental COM quality refactor

Separated Engine declarations/Automation glue, donor operations, class factory,
server lifetime and loader entry point. Research routing now lives under
experiments with shared entry-point declarations. Added stable DispatchId enum,
named custom HRESULTs, ownership/STA comments and formatting/contribution rules.
No private Office calls or IAT behavior changed. Production memory capabilities
remain false. Error-message allocation is now contained at the COM boundary;
EXCEPINFO strings have scoped ownership. Successful public dispatch calls skip
unused diagnostic-message allocation; no speedup is claimed without remeasurement.

Validation: Release and Debug builds pass. Native COM contract tests pass in both
configurations (wrong-thread guard, activation/lifetime, server locks, null
outputs, error HRESULTs). Release PowerPoint fallback-contract, COM smoke and
memory tests pass, including wrong shape types, stale/deleted donors, double
release, missing arguments, repeated Clear, unchanged capabilities, editable
freeforms and zero pixel differences before/after save/reopen. The 100,000-call
stress run was not repeated for this structural change.

The observational GFX probe now names the build-specific factory/stream RVAs,
entry signatures, underlying-stream member and x64 stack argument offsets beside
their evidence. Module snapshots include file versions and the probe rejects an
unknown build before setting a breakpoint. A fresh factory entry/return trace
reproduced the known result and detached successfully; UserPicture completed.
Evidence: docs/evidence/cached_factory_refactor.txt. Older instrumentation and
image-validation formatting remain scheduled for subsequent bounded refactors.

## 2026-09-08 — Initial discovery and baseline

Office: PowerPoint x64 16.0.14334.20848, ProPlus2021Volume Click-to-Run.

Hypothesis: public formatting transfer can reuse an image fill without replacing the destination.

Method: GCC/UCRT64 Release native COM harness, deterministic PNGs, real freeform, capture name/ID/type/position/size/rotation/Z order/node count, call UserPicture then PickUp/Apply, edit a node, change donor and target line weights independently.

Result: identity/geometry snapshot remained equal for both methods. Fill type is 6 (picture). Node SetPosition succeeds after applying. Donor line weight 7 replaces target line weight 1: Apply is whole-style transfer, not fill-only.

Conclusion: useful explicit fallback for compatible styles; cannot silently present it as a fill-only resource binding API. No inference about decoded-image ownership from this alone.

Initial external-process benchmark completed 100/1,000/10,000 fill-operation distributions but received RPC_E_CALL_REJECTED (0x80010001) during deletion after creating 1,010 shapes (includes warmup). No recorded PowerPoint crash. External timings include marshaling and UI scheduling; do not compare these to an in-process optimized backend.

Next: in-process execution, package/resource comparison, controlled native IAT trace.

## 2026-09-08 — In-process activation

Built BlipBridge.dll using GCC C++20, static compiler/runtime linkage, x64 PE. Registered per-user COM Automation and an explicitly connected research COM add-in. Engine PID 16520 matched POWERPNT PID 16520. IDTExtensibility2 OnConnection publishes Engine through COMAddIn.Object; benchmark code executes synchronously on the PowerPoint STA. This avoids changing VBA security or trusting programmatic VBA project access.

Normal automation supports an explicit donor-shape fallback. Byte loading remains E_NOTIMPL until a validated memory pipeline exists. Handles are monotonically increasing integers owning COM references, not exposed native pointers. Release invalidates a handle; Clear does not recycle IDs. Donor lifetime and validity still depend on the presentation staying open.

## 2026-09-08 — Typelib inspection

Enumerated the complete live containing typelib, including restricted/hidden flags, DISPIDs, parameter types and vtable offsets. Raw dumps: artifacts/powerpoint_typelib.txt and artifacts/fill_typelib.txt. FillFormat.UserPicture is DISPID 17, BSTR input, flags 0, typeinfo vtable offset 136. FillFormat contains no stream/byte-array image setter. Shape, ShapeRange, Shapes and PictureFormat are included in the dump. Other UserPicture matches must not be confused with Shape.Fill (e.g. chart formatting).

This rules out a direct setter in this typelib, not undocumented interfaces or internal native functions. Office.js setImage is a modern architectural lead, not evidence this LTSC build exposes the same entry point.

Sources: [PickUp documentation](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.pickup), [PowerPoint API 1.8](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-8-requirement-set).

## 2026-09-08 14:15–14:20 local — File/buffer boundary and RAM fill

Hypothesis: once the ordinary file read is satisfied from a supplied buffer, Office can own/materialize that image without needing a persistent source file.

Method: validated temporary IAT hooks, uniquely named texture, live CaptureStackBackTrace, focused objdump disassembly. Observed OART +0x321750 opening a COM-like object through +0x321970 and reading through its vtable +0x40. Real file path: MSO20 +0x11A638 CreateFileW, +0x7AEDE ReadFile, +0x11AF8D CloseHandle. Full read length 11,644 bytes. See userpicture_pipeline.md for the full chain.

Experiment replaced only the exact nonexistent sentinel name on the initiating thread with a memory buffer behind a Win32 API adapter. An event handle is used as a token; no source file is opened for that name. Returned bytes flow through the ordinary Office loader. PNG/JPEG are not decoded by BlipBridge.

Results: 32/64/128/256 PNGs and JPEG accepted. Final PNG freeform export is pixel-identical to UserPicture and to save/reopen export (zero differing pixels). Saved media SHA256 equals original PNG. Name/ID/type/extents/rotation/Z order/node count verified; node editing succeeds. RAM source file remains nonexistent. This establishes an experimental byte-to-existing-fill path; it does not establish a direct internal-stream API or cached decoded-resource ownership.

An exact-name hardening change initially failed because std::filesystem retained forward slashes while Office normalized them. No match meant ordinary UserPicture rejected the nonexistent path; hooks were restored. Normalizing the sentinel to preferred Windows separators fixed the regression and the full memory functional test passed again.

## 2026-09-08 14:20–14:25 local — Symbols and narrowed candidates

Configured DbgHelp with the Microsoft public symbol server. Local symbol-server helper loading produced error 126. Independently parsed each PE's RSDS identity and requested the exact public PDB URL. Both returned HTTP 404:

- Mso20Win32Client.pdb / 2BEAFC7A15824BE193F1ACEB60A9893B2.
- oart.pdb / 4A62D913850B42DBA52D8F70A8FBF8472.

No private symbol names are inferred. OART +0x89C860 constructs fill/property-update state, calls the filename loader +0x950070, then invokes a document-related vtable method. Stack-backed structures, exception cleanup and custom reference wrappers make a guessed direct call unsafe. No call was attempted. Next: dynamic buffer-consumption trace to find a narrower boundary with understood ownership.

Hidden-document comparison used live typelib IID and function offsets to bind public dual COM calls. UserPicture ~1.09 ms, pre-picked Apply ~0.87 ms, PickUp+Apply ~1.25 ms. DispID resolution is not the dominant cost in these measurements. Donor transfer is not automatically faster than filename loading.

## 2026-09-08, continuation through 19:25 local - correction and standard stream

Office build remains 16.0.14334.20848 x64. Earlier source-only tracing was insufficient: adding WriteFile and Content.MSO handle tracking proved Office writes a real temporary PNG and subsequently reads it through MSO20/GFX/WIC, including during MemoryFillExperiment. The earlier RAM-fill result is source delivery only, not the strict milestone 4. Documentation and capabilities now explicitly retain that limitation. Current validation decodes PNG/JPEG with WIC before invoking Office because malformed image bytes could otherwise return apparent success; the old no-BlipBridge-decode statement applies only to the earlier experiment.

Hypothesis: GFX's decoder wrapper contains a standard IStream despite the earlier private file-copy interface. Method: new read-only GDB object probe with signature validation at GFX +0x7880, live vtable capture, followed by static QueryInterface inspection. Result: underlying MSO20 vtable +0x45F630, QI thunk +0x185820 -> +0x70510, which explicitly recognizes IID_IStream and returns this subobject. UserPicture completed after debugger detach. No inferior function calls or retained pointers. Full RVAs and stack are in userpicture_pipeline.md and evidence/decoder_stream.txt.

Export inspection found actual GEL::ICachedImage::Create(IStream*, ...) at GFX +0x7680 on the observed path and GEL::IImage::Create(IStream*, bool) at +0x194090. These private C++ APIs return Ofc::TCntPtr objects; ABI/lifetime and fill binding are still unresolved. Next experiment should observe caller arguments, returned resource and matching release on the existing Office call path before implementing a stream-factory call. Direct decoder substitution by itself would not remove the earlier cache-file write.

Separately, 100 cached PickUp/Apply calls had zero monitored open/read/write calls on the initiating STA. 100,000 alternating donor applications completed; memory accumulated substantially with document history. Closing the document returned private bytes to about 225 MB. Reopened stress.pptx contains 133 normal image-filled shapes and two shared media resources. This proves serialization/reopen for this workload, not independent internal resource ownership or universal absence of decoding.

Release and Debug builds passed with GCC/UCRT64. Memory functional and COM smoke tests passed again, including malformed-image rejection, stale handles and deleted donors. No native crash observed. A first buffer watch that continued beyond RtlFreeHeap was discarded because later allocator reuse was unrelated. String-anchor unwind-region starts are now labeled accurately; cold fragments are not assumed callable entries.

19:29 local: extended the debugger probe to observe existing ICachedImage::Create entry and return. It returned distinct cached-image and image objects, with vtables GFX +0x409DC0 and +0x4055C8. Static AddRef/Release inspection reveals intrusive count at +8, increment slot 0, decrement slot +8; these must not be cast to IUnknown. Observed caller returns to OART +0x8F554, then consumes the cached result through +0x8F94C and releases a local reference. Evidence: cached_factory.txt. No Office pointers were retained or modified. The first fresh-process attempt had no GFX in its lazy module snapshot; adding an ordinary warmup fill fixed the harness and the subsequent trace completed. Next: determine whether this object is only a rendering cache or can participate in document resource/fill binding without Content.MSO storage.
