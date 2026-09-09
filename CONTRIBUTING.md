# Engineering standards

BlipBridge is native systems research hosted inside PowerPoint. Readability is a
correctness requirement, especially around ABI assumptions and resource lifetime.

- Use descriptive names, ordinary whitespace, braces, and one statement per line.
  The repository `.clang-format` records the intended C++ formatting style.
- Keep COM dispatch limited to argument/result conversion and error translation.
  Separate Office operations, server lifetime, and experimental instrumentation.
- Document ownership, STA requirements, invariants, and unusual ABI assumptions
  beside the implementation. Use concise Doxygen comments on important interfaces.
- Name every private offset/RVA. Record module, exact build, signature, evidence,
  and whether the address is observational or validated for invocation. A matching
  signature alone never establishes a callable ABI or resource ownership.
- Use RAII for COM references, handles, streams, exception strings, and patches.
  No C++ exception may escape a COM or C entry point, including error reporting.
- Keep experimental code under `experiments/`. Explicitly document research code
  linked into the research DLL. Normal LoadTexture/SetImageBytes must not route
  through a file adapter or advertise unvalidated capabilities.
- Preserve hot-path behavior during structural refactoring. Avoid adding allocation,
  name resolution, logging, or ownership churn. Benchmark changed performance claims.
- Refactor one subsystem at a time. Build Release and Debug, run CTest for both,
  and run affected PowerPoint regressions after meaningful changes. Do not rerun a
  100,000-operation stress workload unless the change warrants it.
- Maintain architecture, status, research, pipeline, and compatibility documents.
  Explicitly correct superseded conclusions; separate observations from hypotheses.

Current regression commands, after `build.ps1`:

```powershell
& C:\msys64\ucrt64\bin\ctest.exe --test-dir build/Release --output-on-failure
& C:\msys64\ucrt64\bin\ctest.exe --test-dir build/Debug --output-on-failure
.\tests\fallback_contract.ps1
.\tests\com_smoke.ps1
.\tools\test_memory.ps1
```

The CTest command path above is this machine's discovered toolchain; use the
discovered CMake installation's sibling `ctest` on another machine. PowerPoint
tests require the explicitly registered research add-in and generated textures.
