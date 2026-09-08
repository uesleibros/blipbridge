# Architecture

The compatibility interface is an apartment-threaded in-process IDispatch COM class, BlipBridge.Engine. VBA CreateObject loads it into the VBA host. Creating it from PowerShell instead loads it into PowerShell; use the research COM add-in to invoke benchmarks inside PowerPoint. GetHostProcessId makes this verifiable.

Current implementation has one explicit public-COM fallback. RegisterTextureShape retains a normal AutoShape/Freeform donor whose Fill.Type is picture. ApplyTexture calls donor.PickUp and destination.Apply, preserving destination geometry but replacing broader formatting. It does not create a Picture Shape. It neither reads an image file nor processes PNG bytes during ApplyTexture; Office's internal decoding/resource behavior remains unproven.

The fallback has no byte-loading API implementation: the caller must preload a donor by an existing supported mechanism. This is not milestone 4. MemoryStream and InternalBlip backends remain research targets and are not advertised as working.

No raw Office internal pointer is retained. Handles are per-engine positive integers and never reused during an engine lifetime. COM references own donor wrappers, not a guaranteed independently retained document resource. Release/Clear should precede presentation close. Cross-thread calls are rejected; C++ exceptions are translated at COM entry points.

Research instrumentation is under experiments. IAT tracing validates loaded x64 PE imports and the exact existing imported function pointer, temporarily changes only matching Office import slots, captures stacks, and restores the slots. It is invoked explicitly, never by normal ApplyTexture. It does not patch Office files on disk.
