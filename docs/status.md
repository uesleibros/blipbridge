# Current status

Current backend: **PickupApplyFallback**, experimental, whole-style transfer from an existing normal donor shape with a preloaded picture fill.

Milestone 1: GCC/UCRT64 x64 Release executable and Automation DLL built. Office exact build detected. Reproducible environment probe records loaded modules.

Milestone 2: functional checks passed; external benchmark partially completed; full in-process benchmark underway.

Milestones 3–4: native pipeline investigation underway. No validated memory-byte-to-fill backend yet. LoadTexture and SetImageBytes explicitly return E_NOTIMPL.

What works: retaining donor shapes through safe COM references; fallback ApplyTexture; geometry/identity preserved in the first freeform test; editable nodes after transfer; DLL runs in POWERPNT through research add-in.

What does not work: direct PNG/JPEG bytes, decoded BLIP ownership, fill-only transfer, internal resource cache. No unsupported function pointers are called.

Known errors: external COM benchmark RPC_E_CALL_REJECTED during shape churn. No confirmed native crashes to date.

Next experiment: validated, reversible Office IAT instrumentation for uniquely named UserPicture file; capture actual CreateFileW/ReadFile/CloseHandle stacks and module RVAs.
