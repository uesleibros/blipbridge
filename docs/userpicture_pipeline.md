# UserPicture native pipeline

Latest ownership investigation: [resource_lifetime.md](resource_lifetime.md)
records the OART +0xF0 cached-reference transfer, later record copies and the
phase-confirmed final release during presentation close. The earlier question
whether the cached object survives the loader is now answered affirmatively for
this experiment; a standalone fill-binding ABI remains unidentified.

Probe maintenance, 2026-09-09: the GFX decoder/factory script now uses named RVAs,
layout offsets and signatures with local evidence comments, and checks the exact
GFX file version from the module snapshot. A repeated factory trace after the COM
refactor produced the same object-vtable/caller results. These remain observational
breakpoints; no private function invocation or ownership guarantee is implied.

Build: PowerPoint / OART 16.0.14334.20848 x64. Reproducible experiment: tools/trace_userpicture.ps1, experiments/exp_userpicture/trace.cpp. Raw stacks: artifacts/userpicture_trace.txt.

Dynamic trace uses temporary validated IAT replacements for CreateFileW/ReadFile/CloseHandle in loaded Office modules. Only the uniquely named texture is recorded. All import slots were restored after the call. GDB is installed and was also attached/detached to sample the benchmark. Export-based debugger names such as PPMain are nearest-export labels, not reliable names for individual internal functions.

Observed caller chain (return-address RVAs, not all function entry points):

```text
PowerPoint FillFormat.UserPicture (DISPID 17)
ppcore.dll + 0x8249D8
oart.dll + 0x8A148D
oart.dll + 0x89C9F6
oart.dll + 0x950100
oart.dll + 0x94D81F
oart.dll + 0x9343D8
oart.dll + 0x9342B1
oart.dll + 0x3216A1
oart.dll + 0x32179D
oart.dll + 0x321F6A
mso20win32client.dll + 0x11A0EC
mso20win32client.dll + 0x11A4CD
mso20win32client.dll + 0x11A638
CreateFileW
```

The open used GENERIC_READ and flags 0x40100000. One ReadFile requested the entire 11,644-byte PNG:

```text
oart.dll + 0x321809
mso20win32client.dll + 0x209F5D
mso20win32client.dll + 0x7AEDE
ReadFile
```

Close occurred through OART + 0x321952 and MSO20 + 0x8763D / 0x1645B3 / 0x11AD34 / 0x11AF8D.

Static disassembly corroborates a function entry OART + 0x321750 calling OART + 0x321970 to obtain a COM-like object, then calling that object's vtable slot +0x40 at OART + 0x321803. The next instruction is the observed read return address 0x321809. This object has Release at slot +0x10, but its +0x40 read-related method is not standard IStream.Read. Do not pass IStream to it based on this resemblance. ABI, interface identity and ownership are not yet recovered.

OART + 0x321970 contains path/resource setup; a later factory call returns the object consumed by +0x321750. Raw focused disassembly is retained under artifacts. Decoder creation and final resource binding have not yet been isolated; this is a confirmed file/buffer boundary, not a complete decoded-BLIP call graph.

## Memory-backed API experiment

tools/test_memory.ps1 supplies a Byte array to the in-process experimental method. The experiment calls the existing UserPicture pipeline with a deliberately nonexistent sentinel filename. For that exact name on the initiating STA only, temporary IAT adapters return an event handle used strictly as a token, report byte-array size/file metadata, and satisfy ReadFile from the in-memory compressed buffer. The adapter creates no source image file or Picture shape. However, expanded WriteFile tracing proved that Office itself writes a temporary PNG under INetCache/Content.MSO, then reads it for decoding. This experiment does NOT satisfy the strict no-temporary-image requirement. Normal unrelated file calls delegate unchanged.

First result: PNG 11,822 bytes read at offset zero from RAM; UserPicture succeeded. Original freeform ID/type/geometry/node count stayed equal; node editing succeeded afterward. Saved package contains one normal freeform, no p:pic, and a PNG with SHA256 identical to the input. Exported shape visibly shows the intended deterministic colored texture.

This is **MemoryFileAdapterExperiment**, not InternalStream or InternalBlip. It temporarily instruments a broad import set and is not part of normal SetImageBytes/LoadTexture. It proves source-byte delivery and correct serialization, but retains Office cache-file I/O and materialization. Current input validation also decodes through WIC before Office to reject malformed images; historical adapter timings predate that validation.

## Expanded cache and decoder trace

Tracing now includes WriteFile and Content.MSO handles, plus imports in WindowsCodecs/GDI+. A PNG-bearing WriteFile resolves to a real Content.MSO PNG during both ordinary UserPicture and MemoryFillExperiment. The initial narrow source-filename trace missed this. The source and destination objects share MSO20 vtable RVA 0x475910; its slot +0x40 is a custom copy routine, not standard IStream.Read. The sink is called through slot +0x20 at MSO20 +0x209FAB. No compatible ownership/ABI has been established for calling these objects ourselves.

Observed cache-read and decoder return-address chain (intermediate frames omitted):

```text
OART +0x94D865 -> +0x8F554
GFX +0x76A9 -> +0x9390 -> +0xA714 -> +0xB722 -> +0x6FB6
MSO40 +0x21C0E7 -> +0x1F917E -> +0x1F8B56 -> +0x1F8EDD
WindowsCodecs PNG processing
GFX +0x78FA (adapter entry +0x7880)
MSO20 +0x7B18C -> +0x7A966 -> +0x7AEDE
ReadFile(Content.MSO PNG)
```

Focused disassembly of GFX +0x7880 shows an underlying object's Seek (+0x28) and Read (+0x18) slots, consistent with IStream. This is a decoder input lead; substituting it alone would not eliminate the earlier cache write. Final resource creation/binding and lifetime remain unresolved.

A hardware watch of the original compressed buffer ended at its RtlFreeHeap call. An earlier watch continued after free and observed unrelated allocator reuse; that evidence is discarded. The string FPreprocessImageFromStream has OART LEA references +0x2B7BC7/+0x2B7C5D in a cold code fragment that rejoins +0x8F1B2. Its unwind-region start is not a safe callable function entry.

Separately, 100 cached donor PickUp/Apply operations completed with zero monitored opens, reads, or writes on the invoking STA. This supports resource reuse through public COM, but does not prove absence of all in-memory decoding or cover uninstrumented worker activity. See docs/evidence/cached_apply_trace.txt.

## Standard stream confirmed on continuation

tools/probe_decoder_stream.py attached GDB, validated the GFX +0x7880 entry bytes, captured objects at that breakpoint, removed it and detached without calling functions in the inferior. The disposable UserPicture operation completed afterward.

- GFX decoder adapter vtable: +0x217D10; Read slot +0x18 points to +0x7880.
- Underlying object at adapter +0x18: MSO20 vtable +0x45F630; Read +0x7B100, Seek +0x7B1E0.
- Its QueryInterface thunk +0x185820 adjusts this by -0x50 and jumps to +0x70510. That routine compares GUIDs at +0x460360 (IUnknown) and +0x460370 (IStream: 0000000C-0000-0000-C000-000000000046), returning the same stream subobject for both. This corroborates standard IStream identity independently of slot resemblance.

Actual GFX exports provide stronger leads than nearest-export stack labels:

- GEL::ICachedImage::Create overload accepting IStream*, IStreamCopyInstruction, optional MD4UID and bool: export ordinal 236, RVA +0x7680. It is in the observed decoder call chain.
- GEL::IImage::Create(IStream*, bool): export ordinal 364, RVA +0x194090. Disassembly uses a hidden return-storage pointer in RCX, stream in RDX and bool in R8, then stores an Office smart-pointer result. This overload was inspected statically, not invoked experimentally.

These are private C++ exports, not public COM setters. Ofc::TCntPtr lifetime, exception ABI, lazy decoding and Shape fill/resource binding must be recovered before any native call is enabled. The main DLL still invokes none of these exports. Raw live-object evidence and export names are archived with the research.

Factory entry/return mode then observed the actual ICachedImage::Create call. RCX is hidden result storage, RDX is the image smart-pointer output reference, R8 the confirmed MSO20 IStream, R9 the copy-instruction enum (observed 0). Stack arguments are an MD4UID pointer and bool (observed 1). Return RAX equals the result-storage address. The returned cached object has GFX vtable +0x409DC0; the image output has vtable +0x4055C8. Both are distinct objects, not the stream or a Shape.

Static inspection identifies common vtable slot 0 at GFX +0x90AA0 as an atomic increment of a count at object +8. Slot +8 decrements that count: cached object +0xA1A0, image +0xA170, with different destructor slots on the last release. These are private intrusive reference-counted interfaces, not IUnknown layouts. Dynamic balanced retention has not yet been tested.

The factory returns directly to OART +0x8F554. Its result is passed to +0x8F94C, followed by a release through cached-object slot +8 at return address +0x8F575. The consumer queries further image metadata and updates an OART object; it is not yet established as a Shape fill setter. This narrows the next ownership/binding investigation. A fresh-process probe initially lacked lazy-loaded GFX in its module snapshot; the harness now warms a separate 32px picture fill before snapshotting, and the repeated probe succeeded.
