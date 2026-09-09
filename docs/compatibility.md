# Compatibility

The 2026-09-09 COM refactor preserves existing registration, DISPIDs, public
return types and Office profile behavior. Native COM contract tests pass for
Release and Debug; PowerPoint integration tests ran against the Release DLL.
Research code remains explicitly linked into this experimental DLL, and normal
memory-image capabilities remain disabled.

Test machine: Windows build 26200.9168 x64; PowerPoint LTSC 2021 x64 16.0.14334.20848.

Only this Office build is tested. Public-COM fallback uses discovered methods and does not call undocumented functions. No internal-function compatibility profiles are enabled. Normal MemoryImageToFill and InternalBackend capabilities are false. The explicit experimental memory adapter checks OART version 16.0.14334.20848, x64 PE architecture and the observed code bytes at OART +0x321750 before installing Win32 API import adapters. It never calls that internal RVA. Every replaced import must equal the resolved real Win32 API pointer. Other builds return unsupported.

Office Click-to-Run virtualizes some shared-module paths. A reported loaded path can be absent from the ordinary filesystem; the physical files are under Office/root/vfs/ProgramFilesCommonX64. Environment reports distinguish missing virtual paths rather than inventing file versions or architectures.

Research build: GCC 16.2.0, MinGW-w64 UCRT, CMake 4.4.2. Microsoft SDK and Visual Studio are installed but not required for the main build. No MSVC ABI-dependent internal calls have been attempted.
