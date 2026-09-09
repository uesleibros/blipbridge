#pragma once
/**
 * @file oart_layout.hpp
 * Shared, guarded access to Office-internal objects behind a PowerPoint
 * FillFormat.
 *
 * Everything here is research-only and specific to one Office build. The rules
 * this module exists to enforce:
 *
 *  - No offset is used until the module version matches and the object's vtable
 *    has been compared against the recorded one.
 *  - No private function is called until the exact bytes at its RVA match the
 *    ones recorded when its ABI was derived.
 *  - Every pointer obtained here is **borrowed**. Nothing in this module takes a
 *    reference, and a resolved receiver must never be cached: a Shape deleted
 *    through public COM still passes every check in this chain.
 *
 * Derivations, evidence and per-function ABI notes are in
 * docs/receiver_lookup.md, docs/record_construction.md and docs/oart_abi.md.
 */

#include <windows.h>
#include <oleauto.h>

#include <cstddef>
#include <cstdint>
#include <string>

namespace bb::oart {

/// The single Office build every offset in this module was derived from.
inline constexpr DWORD kSupportedVersionHigh = 0x00100000;   // 16.0
inline constexpr DWORD kSupportedVersionLow = (14334u << 16) | 20848u;
inline constexpr wchar_t kSupportedVersionText[] = L"16.0.14334.20848";

/**
 * A private entry point that must be byte-verified before it is called.
 * `signature` holds the exact bytes found at `rva` when the ABI was derived, so
 * a shifted or patched build fails before the call rather than during it.
 */
struct GuardedFunction {
    std::uintptr_t rva = 0;
    const char* name = nullptr;
    const std::uint8_t* signature = nullptr;
    std::size_t signatureSize = 0;
};

template <std::size_t N>
constexpr GuardedFunction MakeGuarded(std::uintptr_t rva, const char* name,
                                      const std::uint8_t (&signature)[N]) {
    return GuardedFunction{rva, name, signature, N};
}

/**
 * What a PPCORE automation wrapper's vtable proved to be.
 *
 * PowerPoint's `Shape.Fill` is a thin PPCORE object whose vtable is almost
 * entirely identity thunks: slot N loads `this + innerOffset` and calls that
 * object's vtable at N*8. That shape is what makes the walk safe, and it is what
 * gets verified - rather than one hard-coded vtable RVA, which changes between
 * sibling classes and could change between builds.
 */
struct DelegatingWrapper {
    std::size_t innerOffset = 0;      ///< read out of the thunks, not assumed
    unsigned identityThunks = 0;      ///< how many slots delegate that way
    std::uintptr_t vtableRva = 0;     ///< for diagnostics only
};

/// Borrowed Office objects for one Shape. Valid only for the current call.
struct FillTarget {
    std::uintptr_t oartBase = 0;
    std::uintptr_t ppcoreBase = 0;
    void* publicFill = nullptr;   ///< PPCORE FillFormat, the argument itself
    void* handler = nullptr;      ///< OART FillFormat at publicFill + wrapper.innerOffset
    void* token = nullptr;        ///< control block at handler+0x58
    void* receiver = nullptr;     ///< OART receiver at token+0x10
    std::uint32_t tokenStrong = 0;
    std::uint8_t handlerFlag = 0; ///< handler+0x60, the transaction's bool argument
    DelegatingWrapper wrapper;    ///< how the public object was recognised
};

/// Version text of a loaded module, or "not loaded"/"unreadable" - diagnostics only.
std::wstring ModuleVersionText(const wchar_t* moduleName);

/**
 * Examines the vtable of @p object and reports the delegating shape it has, if
 * any. Returns false when the object is not a delegating wrapper at all, which
 * is the normal answer for, say, a Shape passed where a FillFormat was meant.
 */
bool DescribeDelegatingWrapper(const void* object, std::uintptr_t moduleBase,
                               std::size_t moduleSize, DelegatingWrapper& wrapper);

/// True when every page spanning [address, address+size) is committed and readable.
bool IsReadable(const void* address, std::size_t size);

std::uintptr_t LoadPointer(const void* base, std::size_t offset);
std::uint32_t LoadDword(const void* base, std::size_t offset);

/// Describes an address as `module+0xRVA`, so a failed check is diagnosable.
std::string DescribeAddress(std::uintptr_t address);

/**
 * Returns the named Office module, resolving it from OART's directory if it is
 * not loaded yet. GFX is delay-loaded, so it is absent until Office's first
 * picture operation; this makes the backend usable without a warm-up fill.
 */
HMODULE EnsureOfficeModule(const wchar_t* moduleName);

/// Throws bb::Error unless the named module is loaded and is the supported build.
HMODULE RequireSupportedModule(const wchar_t* moduleName, const char* description);

/// The three Office modules an apply depends on, as one lookup.
struct OfficeModules {
    HMODULE oart = nullptr;
    HMODULE ppcore = nullptr;
    /// Delay-loaded: legitimately null until Office's first picture operation.
    HMODULE gfx = nullptr;
};

/**
 * Fetches all three module handles once and points the validation cache at them,
 * discarding everything it holds if any of them changed.
 *
 * This exists because the handles are the cache's *reload detector*, so they
 * cannot themselves be cached - but fetching them once per operation rather than
 * three times per module check costs a tenth as much and detects exactly the
 * same thing. GFX is looked up and never loaded here: only the create path needs
 * it, and loading it as a side effect of an apply would be a surprise.
 */
OfficeModules SynchroniseOfficeModules();

/**
 * Version-checks an already-resolved module handle. Identical to
 * RequireSupportedModule except that the caller has done the lookup, so nothing
 * is fetched twice.
 */
HMODULE RequireValidatedModule(HMODULE module, const char* description);

/**
 * The module's SizeOfImage, read once per loaded image and then remembered.
 * A property of the image rather than of any document, so it is cached under the
 * same rule as everything else in the validation cache.
 */
std::size_t OfficeModuleImageSize(HMODULE module);

/**
 * Walks Shape.Fill to the OART receiver, verifying both module versions and
 * every vtable on the way. Throws bb::Error naming the check that failed.
 *
 * The result is borrowed and must be re-resolved for every operation; see the
 * caching hazard in docs/receiver_lookup.md.
 */
FillTarget ResolveFillTarget(IDispatch* fill);

/// Throws bb::Error unless the bytes at `moduleBase + function.rva` still match.
void RequireSignature(std::uintptr_t moduleBase, const GuardedFunction& function);

/// Resolves a verified private entry point to a callable address.
void* GuardedAddress(std::uintptr_t moduleBase, const GuardedFunction& function);

// -- intrusive counting shared by the GFX image objects ---------------------

/// GFX intrusive objects keep a 32-bit count at +8; Release is vtable slot +8.
inline constexpr std::size_t kIntrusiveCountOffset = 8;
inline constexpr std::size_t kIntrusiveReleaseSlot = 8;

std::uint32_t IntrusiveCount(const void* object);

/// Releases one reference. Never read the object afterwards.
void ReleaseIntrusive(void* object) noexcept;

} // namespace bb::oart
