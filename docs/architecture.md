# Architecture

> **Scope.** This page describes the **research COM surface** - the
> `BlipBridge.Engine` Automation object the Office harnesses drive - and dates
> from before the shipped C ABI existed. For the architecture of the library
> people actually install, read [c_abi.md](c_abi.md): it has the seam between the
> C ABI and the two Office backends, which of the two each architecture gets, and
> which modules they share. For what separates the backends at runtime, see
> [compatibility.md](compatibility.md) and [capabilities.md](capabilities.md).
>
> The research surface itself is now built over whichever backend a configuration
> has, which is what lets one harness suite drive both.

The compatibility interface is an apartment-threaded in-process IDispatch COM class, BlipBridge.Engine. VBA CreateObject loads it into the VBA host. Creating it from PowerShell instead loads it into PowerShell; use the research COM add-in to invoke benchmarks inside PowerPoint. GetHostProcessId makes this verifiable.

Current implementation has one explicit public-COM fallback. RegisterTextureShape retains a normal AutoShape/Freeform donor whose Fill.Type is picture. ApplyTexture calls donor.PickUp and destination.Apply, preserving destination geometry but replacing broader formatting. It does not create a Picture Shape. It neither reads an image file nor processes PNG bytes during ApplyTexture; Office's internal decoding/resource behavior remains unproven.

The normal fallback has no byte-loading API implementation: the caller must preload a donor. The separate MemoryFillExperiment delivers Byte arrays through reversible API adapters, but Office subsequently writes and rereads a temporary Content.MSO PNG. It fails the strict file-free requirement. Subsequent registered-handle applications use only public COM without hooks or source bytes; 100 traced applications made zero monitored file calls on the invoking STA. This is a whole-style fallback, not an independently owned decoded-BLIP cache. MemoryStream and InternalBlip remain research targets.

No raw Office internal pointer is retained. Handles are per-engine positive integers and never reused during an engine lifetime. COM references own donor wrappers, not a guaranteed independently retained document resource. Release/Clear should precede presentation close. Cross-thread calls are rejected; C++ exceptions are translated at COM entry points.

Research instrumentation is under experiments. IAT tracing validates loaded x64 PE imports and the exact existing imported function pointer, temporarily changes only matching Office import slots, captures stacks, and restores the slots. It is invoked explicitly, never by normal ApplyTexture. It does not patch Office files on disk.

## COM refactor boundaries

- `src/com/engine.hpp` declares the stable DispatchId enum, borrowed Automation
  argument reader, add-in ABI, and apartment-bound Engine ownership.
- `src/com/engine.cpp` implements Automation conversion, dispatch, and exception
  translation. Texture operations are in `engine_textures.cpp`; the existing map
  still owns donor COM wrappers and never recycles handles.
- `src/com/class_factory.cpp` owns activation and process-wide server reference/
  lock accounting. `src/dllmain.cpp` performs only loader-safe notification setup.
- `experiments/automation_bridge.cpp` explicitly routes research methods. The
  research build still links experimental code into the same DLL, as before;
  this is source separation, not a claim of separate production/research binaries.
- `include/blipbridge/errors.hpp` centralizes current custom HRESULTs. Shared
  Automation Value and exception-string RAII remain in `dispatch.hpp`, used by
  standalone tools as well as the DLL.

The COM refactor preserves CLSID, ProgID, all 18 DISPIDs, return types, handle
semantics, and donor PickUp/Apply ordering. Error reporting now contains allocation
failures; invalid pointer arguments fail explicitly. Successful dispatch-helper
calls no longer construct unused diagnostic strings. No new performance numbers
are claimed from that small hot-path change.
