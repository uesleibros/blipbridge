# Research journal

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
