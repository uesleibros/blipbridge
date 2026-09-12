# Compatibility

## Which backend runs where

BlipBridge has two Office backends behind one public API. Everything below this
section is about the **accelerated** one - the Office build profile it is guarded
to, and the evidence behind each private entry point it uses. None of it applies
to the portable backend, which calls no Office internals at all.

| | accelerated | portable |
|---|---|---|
| how it fills | reverse-engineered OART picture-fill transaction | `Shape.Fill.UserPicture`, `ShapeRange.Fill` |
| Office builds | exactly 16.0.14334.20848, x64; fails closed elsewhere | any build that has the documented API |
| architectures | x64 only | x64 (forced) and x86 |
| capability mask | `0x03FF` | `0x03FE` |
| tied to an Office version | yes, by design | no |

The portable backend is therefore the answer to "what happens on an Office build
BlipBridge has not been validated against" as well as to "what happens on 32-bit".
On an unvalidated x64 build the accelerated backend refuses and `BB_ApplyPicture`
still works through Office's own route; a build configured for the portable
backend has no version guard to fail in the first place.

**What is not a fallback.** Backend *selection* happens once, at build time, and
is visible through `BB_CAP_NATIVE_BACKEND`. It is not a runtime failover: the
accelerated backend does not quietly hand work to the portable one when an
internal call fails. A genuine internal failure is reported as itself, because
retrying it on a slower path would turn "your Office build is unsupported" into
"everything is a bit slow" and hide exactly the problems worth knowing about. The
one documented dispatch that does choose a route per call is
`BB_ApplyPicture`, which picks from the Shape's *class* before doing any work -
see picture_cache.md.

The 2026-09-09 COM refactor preserves existing registration, DISPIDs, public
return types and Office profile behavior. Native COM contract tests pass for
Release and Debug; PowerPoint integration tests ran against the Release DLL.
Research code remains explicitly linked into this experimental DLL, and normal
memory-image capabilities remain disabled.

Test machine: Windows build 26200.9168 x64; PowerPoint LTSC 2021 x64 16.0.14334.20848.

Only this Office build is tested. Public-COM fallback uses discovered methods and does not call undocumented functions. No internal-function compatibility profiles are enabled. Normal MemoryImageToFill and InternalBackend capabilities are false. The explicit experimental memory adapter checks OART version 16.0.14334.20848, x64 PE architecture and the observed code bytes at OART +0x321750 before installing Win32 API import adapters. It never calls that internal RVA. Every replaced import must equal the resolved real Win32 API pointer. Other builds return unsupported.

Office Click-to-Run virtualizes some shared-module paths. A reported loaded path can be absent from the ordinary filesystem; the physical files are under Office/root/vfs/ProgramFilesCommonX64. Environment reports distinguish missing virtual paths rather than inventing file versions or architectures.

Debugger research anchors are validated the same way. `probe_fill_transaction.py`
checks the reported OART version and compares recorded instruction bytes at every
one of its eleven breakpoint addresses before attaching them; a single mismatch
aborts the run and detaches. Those addresses are read-only observation points, not
call targets, so they widen no compatibility profile. The GFX and OART offsets they
depend on (image sub-record +0x90/+0xF0, operation size 0x570, constructor
+0x10244, intrusive count at object+8) are documented in fill_transaction.md and
must be re-derived per build before any of them is used for anything but reading.

The one Office function BlipBridge now calls is an export, not an offset:
`GEL::ICachedImage::Create` from GFX.DLL, resolved with GetProcAddress from its
mangled name. It is still guarded by the exact GFX version and by checking the
returned objects' vtables before touching them, and an unrecognised vtable causes
the pointer to be dropped rather than released. No private OART offset is called.

The native apply experiment additionally calls thirteen private OART entry
points. Each is guarded three ways: oart.dll, ppcore.dll and gfx.dll must all be
16.0.14334.20848; the sixteen bytes at each function's RVA must match the ones
recorded when its ABI was derived; and every object walked to must present its
recorded vtable. Any single mismatch aborts before an object is constructed. The
full table, with ownership and destructor pairing, is in docs/oart_abi.md. This
remains research-only: no capability is enabled by it.

Research build: GCC 16.2.0, MinGW-w64 UCRT, CMake 4.4.2. Microsoft SDK and Visual Studio are installed but not required for the main build. No MSVC ABI-dependent internal calls have been attempted.
