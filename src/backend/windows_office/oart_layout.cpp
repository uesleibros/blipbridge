/**
 * @file oart_layout.cpp
 * Implementation of the guarded Office-internal accessors.
 *
 * The offsets below are observational on Office 16.0.14334.20848 x64. Their
 * derivation is in docs/receiver_lookup.md:
 *
 *  - PowerPoint's `Shape.Fill` is a thin PPCORE wrapper whose vtable is almost
 *    entirely identity thunks: slot N loads `this + innerOffset` and calls the
 *    inner object's vtable at N*8. `PPCORE +0x8249C0` is slot 17 of one such
 *    table and forwards `this+0x08` to inner vtable +0x88, which is where the
 *    OART FillFormat vtable holds the UserPicture entry.
 *
 *    That *shape* is what gets verified, not one vtable address. PPCORE has a
 *    whole family of these wrappers - `Shape.Fill`, `ShapeRange.Fill`,
 *    `Shape.Line`, `Shape.TextFrame` all have their own - and pinning one RVA
 *    rejects the siblings and would break on any build that moves the table.
 *    The inner offset is read out of the thunks rather than assumed, and the
 *    inner object still has to present the exact OART FillFormat vtable, which
 *    is what separates a Fill from a Line.
 *  - The OART FillFormat's vtable OART +0xAF60B8 is an IDispatch layout with a
 *    non-atomic uint32 reference count at +0x30, which is why this whole chain
 *    is STA-only.
 *  - Its factory OART +0x23B220 AddRefs a control block into +0x58; OART
 *    +0x63EA0 is a sixteen-byte function that returns `*(token + 0x10)`.
 *  - The receiver's constructor OART +0x223950 installs vtable +0x9F6658.
 */

#include "oart_layout.hpp"

#include <blipbridge/dispatch.hpp>
#include <cstring>
#include <map>
#include <psapi.h>
#include <set>
#include <sstream>
#include <vector>

namespace bb::oart {
namespace {

/// The wrapper vtable observed for Shape.Fill on the tested build; diagnostics only.
constexpr std::uintptr_t kObservedFillFormatVtableRva = 0x1464478; // ppcore.dll
constexpr std::size_t kPublicFillFormatSize = 0x10;
/// How many identity thunks a vtable must have before it counts as a wrapper.
/// The observed Shape.Fill table has 23; the sibling wrappers have 10 to 31.
constexpr unsigned kMinimumIdentityThunks = 6;
/// How far into a vtable to look. The observed tables delegate within this range.
constexpr unsigned kWrapperSlotsExamined = 48;

constexpr std::size_t kFillFormatSize = 0x68;
constexpr std::size_t kFillFormatTokenOffset = 0x58;
constexpr std::size_t kFillFormatFlagOffset = 0x60;

constexpr std::size_t kControlBlockSize = 0x18;
constexpr std::size_t kControlBlockStrongOffset = 0x00;
constexpr std::size_t kControlBlockPointeeOffset = 0x10;

constexpr std::size_t kReceiverInspectedSize = 0x28;

} // namespace

namespace {

/**
 * Decodes one identity thunk, if that is what @p code is.
 *
 * The shape, as emitted for every delegating slot:
 *
 *     48 83 ec ??             sub  $imm8,%rsp          (optional)
 *     48 8b 49 XX             mov  XX(%rcx),%rcx       inner = this->at_XX
 *     48 8b 01                mov  (%rcx),%rax         its vtable
 *     48 8b 80 YY YY YY YY    mov  YY(%rax),%rax       the slot   (or 48 8b 40 YY)
 *     ff 15 ...               call *disp32(%rip)
 *
 * Returns false for anything else, including a real method implementation.
 */
bool DecodeIdentityThunk(const std::uint8_t* code,
                         std::size_t& innerOffset,
                         std::size_t& slotOffset) {
    std::size_t cursor = 0;
    if (code[0] == 0x48 && code[1] == 0x83 && code[2] == 0xEC) {
        cursor = 4;
    }
    if (code[cursor] != 0x48 || code[cursor + 1] != 0x8B || code[cursor + 2] != 0x49) {
        return false;
    }
    innerOffset = code[cursor + 3];
    cursor += 4;
    if (code[cursor] != 0x48 || code[cursor + 1] != 0x8B || code[cursor + 2] != 0x01) {
        return false;
    }
    cursor += 3;
    if (code[cursor] == 0x48 && code[cursor + 1] == 0x8B && code[cursor + 2] == 0x80) {
        std::uint32_t value = 0;
        std::memcpy(&value, code + cursor + 3, sizeof(value));
        slotOffset = value;
        cursor += 7;
    } else if (code[cursor] == 0x48 && code[cursor + 1] == 0x8B && code[cursor + 2] == 0x40) {
        slotOffset = code[cursor + 3];
        cursor += 4;
    } else {
        return false;
    }
    return code[cursor] == 0xFF && code[cursor + 1] == 0x15;
}

} // namespace

namespace {

/**
 * Caches what cannot change while a module stays loaded.
 *
 * Validating a module means reading its version resource off disk, and
 * validating a wrapper vtable means decoding up to 48 thunks with a VirtualQuery
 * apiece. Doing either per apply made the hot path several times more expensive
 * than the work it guards, which the benchmark caught.
 *
 * What is safe to cache is exactly what is a property of the loaded image: the
 * module's identity and base, the shape of a vtable at a given address, and the
 * bytes at a function's RVA. The cache is keyed by the module handle and thrown
 * away whenever that handle changes, so an unload/reload re-validates from
 * scratch.
 *
 * What is deliberately **not** cached is anything derived from a document: the
 * FillFormat, the control block and the receiver are re-read and re-checked on
 * every single call, because a deleted Shape still passes every one of those
 * checks and a stale receiver would be a use-after-free.
 */
class ValidationCache {
  public:
    static ValidationCache& Instance() {
        static ValidationCache cache;
        return cache;
    }

    /// True when this exact module handle has already been version-checked.
    bool IsModuleValidated(HMODULE module) const {
        return validatedModules_.find(module) != validatedModules_.end();
    }

    void RememberModule(HMODULE module) {
        validatedModules_.insert(module);
    }

    /// Looks up a previously validated vtable shape, keyed by its address.
    const DelegatingWrapper* FindWrapper(std::uintptr_t vtable) const {
        const auto found = wrappers_.find(vtable);
        return found == wrappers_.end() ? nullptr : &found->second;
    }

    void RememberWrapper(std::uintptr_t vtable, const DelegatingWrapper& wrapper) {
        wrappers_.emplace(vtable, wrapper);
    }

    /**
     * SizeOfImage for a loaded module. GetModuleInformation measured ten
     * microseconds a call - on its own the most expensive primitive in the
     * guard chain - and the answer cannot change while the handle is the same.
     */
    std::size_t ImageSize(HMODULE module) {
        const auto found = imageSizes_.find(module);
        if (found != imageSizes_.end()) {
            return found->second;
        }
        MODULEINFO information{};
        if (!GetModuleInformation(GetCurrentProcess(), module, &information, sizeof(information))) {
            return 0;
        }
        imageSizes_.emplace(module, information.SizeOfImage);
        return information.SizeOfImage;
    }

    /// Drops everything if any module of interest was unloaded or moved.
    void SynchroniseWith(HMODULE oart, HMODULE ppcore, HMODULE gfx) {
        if (oart == oart_ && ppcore == ppcore_ && gfx == gfx_) {
            return;
        }
        validatedModules_.clear();
        wrappers_.clear();
        imageSizes_.clear();
        oart_ = oart;
        ppcore_ = ppcore;
        gfx_ = gfx;
    }

  private:
    ValidationCache() = default;

    std::set<HMODULE> validatedModules_;
    std::map<std::uintptr_t, DelegatingWrapper> wrappers_;
    std::map<HMODULE, std::size_t> imageSizes_;
    HMODULE oart_ = nullptr;
    HMODULE ppcore_ = nullptr;
    HMODULE gfx_ = nullptr;
};

} // namespace

std::wstring ModuleVersionText(const wchar_t* moduleName) {
    HMODULE module = GetModuleHandleW(moduleName);
    if (!module) {
        return L"not loaded";
    }
    wchar_t modulePath[MAX_PATH * 4]{};
    if (!GetModuleFileNameW(module, modulePath, static_cast<DWORD>(std::size(modulePath)))) {
        return L"unreadable";
    }
    DWORD ignored = 0;
    const DWORD versionSize = GetFileVersionInfoSizeW(modulePath, &ignored);
    std::vector<BYTE> version(versionSize);
    if (!versionSize || !GetFileVersionInfoW(modulePath, 0, versionSize, version.data())) {
        return L"unreadable";
    }
    VS_FIXEDFILEINFO* information = nullptr;
    UINT length = 0;
    if (!VerQueryValueW(version.data(), L"\\", reinterpret_cast<void**>(&information), &length) ||
        length < sizeof(*information)) {
        return L"unreadable";
    }
    std::wostringstream out;
    out << HIWORD(information->dwFileVersionMS) << L'.' << LOWORD(information->dwFileVersionMS)
        << L'.' << HIWORD(information->dwFileVersionLS) << L'.'
        << LOWORD(information->dwFileVersionLS);
    return out.str();
}

bool DescribeDelegatingWrapper(const void* object,
                               std::uintptr_t moduleBase,
                               std::size_t moduleSize,
                               DelegatingWrapper& wrapper) {
    if (!IsReadable(object, sizeof(void*))) {
        return false;
    }
    const std::uintptr_t vtable = LoadPointer(object, 0);
    if (vtable < moduleBase || vtable >= moduleBase + moduleSize) {
        return false;
    }
    if (const DelegatingWrapper* known = ValidationCache::Instance().FindWrapper(vtable)) {
        // A vtable's shape is a property of the loaded image, not of any object,
        // so one analysis per address is enough.
        wrapper = *known;
        return true;
    }
    wrapper = DelegatingWrapper{};
    wrapper.vtableRva = vtable - moduleBase;
    if (!IsReadable(reinterpret_cast<const void*>(vtable), kWrapperSlotsExamined * sizeof(void*))) {
        return false;
    }

    // Every thunk must agree on the inner offset; a table where they disagree is
    // not the simple wrapper this walk depends on.
    bool haveOffset = false;
    for (unsigned slot = 0; slot < kWrapperSlotsExamined; ++slot) {
        const std::uintptr_t entry =
            LoadPointer(reinterpret_cast<const void*>(vtable), slot * sizeof(void*));
        if (entry < moduleBase || entry >= moduleBase + moduleSize) {
            continue;
        }
        constexpr std::size_t kThunkBytes = 24;
        if (!IsReadable(reinterpret_cast<const void*>(entry), kThunkBytes)) {
            continue;
        }
        std::size_t innerOffset = 0;
        std::size_t slotOffset = 0;
        if (!DecodeIdentityThunk(
                reinterpret_cast<const std::uint8_t*>(entry), innerOffset, slotOffset)) {
            continue;
        }
        if (slotOffset != slot * sizeof(void*)) {
            continue; // delegates, but not to the matching slot
        }
        if (haveOffset && innerOffset != wrapper.innerOffset) {
            return false;
        }
        wrapper.innerOffset = innerOffset;
        haveOffset = true;
        ++wrapper.identityThunks;
    }
    if (!haveOffset || wrapper.identityThunks < kMinimumIdentityThunks) {
        return false; // not a wrapper; nothing worth remembering
    }
    ValidationCache::Instance().RememberWrapper(vtable, wrapper);
    return true;
}

bool IsReadable(const void* address, std::size_t size) {
    if (!address || size == 0) {
        return false;
    }
    auto cursor = reinterpret_cast<const std::uint8_t*>(address);
    const auto* end = cursor + size;
    while (cursor < end) {
        MEMORY_BASIC_INFORMATION information{};
        if (VirtualQuery(cursor, &information, sizeof(information)) != sizeof(information)) {
            return false;
        }
        if (information.State != MEM_COMMIT) {
            return false;
        }
        constexpr DWORD kReadableMask = PAGE_READONLY | PAGE_READWRITE | PAGE_WRITECOPY |
                                        PAGE_EXECUTE_READ | PAGE_EXECUTE_READWRITE |
                                        PAGE_EXECUTE_WRITECOPY;
        if (!(information.Protect & kReadableMask) ||
            (information.Protect & (PAGE_GUARD | PAGE_NOACCESS))) {
            return false;
        }
        cursor =
            reinterpret_cast<const std::uint8_t*>(information.BaseAddress) + information.RegionSize;
    }
    return true;
}

std::uintptr_t LoadPointer(const void* base, std::size_t offset) {
    std::uintptr_t value = 0;
    std::memcpy(&value, reinterpret_cast<const std::uint8_t*>(base) + offset, sizeof(value));
    return value;
}

std::uint32_t LoadDword(const void* base, std::size_t offset) {
    std::uint32_t value = 0;
    std::memcpy(&value, reinterpret_cast<const std::uint8_t*>(base) + offset, sizeof(value));
    return value;
}

std::string DescribeAddress(std::uintptr_t address) {
    HMODULE owner = nullptr;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                            reinterpret_cast<LPCWSTR>(address),
                            &owner) ||
        !owner) {
        std::ostringstream out;
        out << "0x" << std::hex << address;
        return out.str();
    }
    char modulePath[MAX_PATH]{};
    GetModuleFileNameA(owner, modulePath, static_cast<DWORD>(std::size(modulePath)));
    const char* name = std::strrchr(modulePath, '\\');
    std::ostringstream out;
    out << (name ? name + 1 : modulePath) << "+0x" << std::hex
        << (address - reinterpret_cast<std::uintptr_t>(owner));
    return out.str();
}

HMODULE EnsureOfficeModule(const wchar_t* moduleName) {
    if (HMODULE loaded = GetModuleHandleW(moduleName)) {
        return loaded;
    }
    // GFX is a delay-load dependency of the Office image pipeline, so it is
    // absent until Office does its first picture operation. Resolving it early
    // is exactly what Office's own delay-load thunk would do on first use, and
    // loading it by full path from OART's directory guarantees the same file
    // rather than something else on the search path.
    //
    // The reference is deliberately never released: the module stays for the
    // life of the process, which is also what Office's delay load does.
    HMODULE anchor = GetModuleHandleW(L"oart.dll");
    if (!anchor) {
        return nullptr;
    }
    wchar_t anchorPath[MAX_PATH * 4]{};
    if (!GetModuleFileNameW(anchor, anchorPath, static_cast<DWORD>(std::size(anchorPath)))) {
        return nullptr;
    }
    std::wstring path(anchorPath);
    const std::size_t separator = path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return nullptr;
    }
    path.replace(separator + 1, std::wstring::npos, moduleName);
    return LoadLibraryExW(path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
}

/**
 * Takes a permanent reference on @p module, so its handle can never go stale.
 *
 * Pinning is what makes the fast path below sound rather than merely fast: a
 * pinned module cannot be unloaded, so its base cannot change and a different
 * image cannot appear at the same address. Windows offers no way to undo it,
 * which is the point.
 */
HMODULE PinModule(HMODULE module) {
    if (!module) {
        return nullptr;
    }
    HMODULE pinned = nullptr;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_PIN |
                                GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                            reinterpret_cast<LPCWSTR>(module),
                            &pinned)) {
        // Not pinnable: keep answering the slow way rather than pretend.
        return nullptr;
    }
    return pinned;
}

OfficeModules SynchroniseOfficeModules() {
    /*
     * Three GetModuleHandleW calls measured 7.1 microseconds together - 38% of
     * the whole receiver resolution, and the largest single cost in it. Each one
     * takes the loader lock and walks Office's module list by name, and it did
     * that on every apply to answer a question whose answer cannot change.
     *
     * Cannot change once the modules are *pinned*, which is what this does. Only
     * an unload could move a module, and a pinned module is never unloaded, so
     * after the first complete lookup the handles are facts about the process
     * rather than a cache of something that might have moved. That is a stronger
     * guarantee than the per-call re-lookup it replaces: the old code could only
     * notice a swap after it happened.
     *
     * Until all three are resolved and pinned, every call does the full lookup.
     * GFX is delay-loaded and is genuinely absent until Office's first picture
     * operation, so "not yet" is an ordinary state, not a failure.
     *
     * STA-only, like the rest of this file: the statics are not synchronised
     * because nothing here may be touched off the owning thread.
     */
    static OfficeModules pinned;
    static bool settled = false;
    if (settled) {
        return pinned;
    }

    OfficeModules modules;
    modules.oart = EnsureOfficeModule(L"oart.dll");
    modules.ppcore = EnsureOfficeModule(L"ppcore.dll");
    modules.gfx = GetModuleHandleW(L"gfx.dll");
    ValidationCache::Instance().SynchroniseWith(modules.oart, modules.ppcore, modules.gfx);

    const OfficeModules candidate{
        PinModule(modules.oart), PinModule(modules.ppcore), PinModule(modules.gfx)};
    if (candidate.oart == modules.oart && candidate.ppcore == modules.ppcore &&
        candidate.gfx == modules.gfx && candidate.oart && candidate.ppcore && candidate.gfx) {
        pinned = candidate;
        settled = true;
    }
    return modules;
}

std::size_t OfficeModuleImageSize(HMODULE module) {
    return module ? ValidationCache::Instance().ImageSize(module) : 0;
}

HMODULE RequireSupportedModule(const wchar_t* moduleName, const char* description) {
    HMODULE module = EnsureOfficeModule(moduleName);
    SynchroniseOfficeModules();
    return RequireValidatedModule(module, description);
}

HMODULE RequireValidatedModule(HMODULE module, const char* description) {
    if (!module) {
        throw bb::Error(E_NOTIMPL,
                        std::string(description) + " is not loaded and could not be resolved");
    }
    ValidationCache& cache = ValidationCache::Instance();
    if (cache.IsModuleValidated(module)) {
        return module; // same loaded image; its version cannot have changed
    }

    wchar_t modulePath[MAX_PATH * 4]{};
    if (!GetModuleFileNameW(module, modulePath, static_cast<DWORD>(std::size(modulePath)))) {
        throw bb::Error(E_NOTIMPL, std::string("Cannot resolve the ") + description + " path");
    }
    DWORD ignored = 0;
    const DWORD versionSize = GetFileVersionInfoSizeW(modulePath, &ignored);
    std::vector<BYTE> version(versionSize);
    if (!versionSize || !GetFileVersionInfoW(modulePath, 0, versionSize, version.data())) {
        throw bb::Error(E_NOTIMPL,
                        std::string("Cannot read the ") + description + " version resource");
    }
    VS_FIXEDFILEINFO* information = nullptr;
    UINT length = 0;
    if (!VerQueryValueW(version.data(), L"\\", reinterpret_cast<void**>(&information), &length) ||
        length < sizeof(*information)) {
        throw bb::Error(E_NOTIMPL,
                        std::string("Cannot read the ") + description + " version information");
    }
    if (information->dwFileVersionMS != kSupportedVersionHigh ||
        information->dwFileVersionLS != kSupportedVersionLow) {
        throw bb::Error(E_NOTIMPL,
                        std::string(description) +
                            " is not Office 16.0.14334.20848; internal layouts are unvalidated");
    }
    cache.RememberModule(module);
    return module;
}

FillTarget ResolveFillTarget(IDispatch* fill) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing FillFormat");
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Office-internal access requires the PowerPoint host");
    }

    // One lookup of all three handles, which also re-points the validation cache
    // and discards it if any module was unloaded or moved. Doing this once per
    // resolve rather than once per module check is the whole difference between
    // a guard chain that costs 37 microseconds and one that costs 12.
    const OfficeModules modules = SynchroniseOfficeModules();

    FillTarget target;
    target.oartBase =
        reinterpret_cast<std::uintptr_t>(RequireValidatedModule(modules.oart, "OART"));
    target.ppcoreBase =
        reinterpret_cast<std::uintptr_t>(RequireValidatedModule(modules.ppcore, "PPCORE"));

    // Step 1: the argument must be a PPCORE delegating wrapper. This is a
    // structural test, not a vtable address: PPCORE has a family of these and
    // pinning one RVA rejects Shape.Fill's own siblings.
    target.publicFill = fill;
    if (!IsReadable(target.publicFill, kPublicFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "FillFormat storage is not readable");
    }
    const std::size_t ppcoreSize = OfficeModuleImageSize(modules.ppcore);
    if (ppcoreSize == 0) {
        throw bb::Error(E_NOTIMPL, "Cannot measure the PPCORE image");
    }
    if (!DescribeDelegatingWrapper(
            target.publicFill, target.ppcoreBase, ppcoreSize, target.wrapper)) {
        std::ostringstream out;
        out << "The object passed is not a PowerPoint automation wrapper: its vtable "
            << DescribeAddress(LoadPointer(target.publicFill, 0))
            << " has no delegating thunks, so it has no inner Office object. "
               "Shape.Fill and ShapeRange.Fill are wrappers; a Shape, ShapeRange or "
               "Slide is not. Pass the FillFormat itself.";
        throw bb::Error(E_NOTIMPL, out.str());
    }

    // Step 2: the OART FillFormat the wrapper forwards to. The inner offset came
    // out of the thunks; this vtable check is what distinguishes a Fill wrapper
    // from the identically shaped Line and TextFrame wrappers.
    target.handler =
        reinterpret_cast<void*>(LoadPointer(target.publicFill, target.wrapper.innerOffset));
    if (!IsReadable(target.handler, kFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "Inner object storage is not readable");
    }
    const std::uintptr_t handlerVtable = LoadPointer(target.handler, 0);
    if (handlerVtable != target.oartBase + kFillFormatVtableRva) {
        std::ostringstream out;
        out << "The wrapper at " << DescribeAddress(LoadPointer(target.publicFill, 0))
            << " delegates through this+0x" << std::hex << target.wrapper.innerOffset << " to "
            << DescribeAddress(handlerVtable) << ", not the validated OART+0x"
            << kFillFormatVtableRva
            << " FillFormat layout. Shape.Line and Shape.TextFrame look like this.";
        throw bb::Error(E_NOTIMPL, out.str());
    }
    target.handlerFlag =
        static_cast<std::uint8_t>(LoadDword(target.handler, kFillFormatFlagOffset) & 0xFF);

    // Step 3: the control block held at +0x58.
    target.token = reinterpret_cast<void*>(LoadPointer(target.handler, kFillFormatTokenOffset));
    if (!IsReadable(target.token, kControlBlockSize)) {
        throw bb::Error(E_NOTIMPL, "Receiver control block is not readable");
    }
    target.tokenStrong = LoadDword(target.token, kControlBlockStrongOffset);

    // Step 4: the receiver, confirmed by its own vtable before any field read.
    target.receiver =
        reinterpret_cast<void*>(LoadPointer(target.token, kControlBlockPointeeOffset));
    if (!IsReadable(target.receiver, kReceiverInspectedSize)) {
        throw bb::Error(E_NOTIMPL, "Receiver storage is not readable");
    }
    const std::uintptr_t receiverVtable = LoadPointer(target.receiver, 0);
    if (receiverVtable != target.oartBase + kReceiverVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "Receiver vtable is " + DescribeAddress(receiverVtable) +
                            ", not the validated OART+0x9F6658 layout");
    }
    return target;
}

void RequireSignature(std::uintptr_t moduleBase, const GuardedFunction& function) {
    const auto* code = reinterpret_cast<const std::uint8_t*>(moduleBase + function.rva);
    if (!IsReadable(code, function.signatureSize)) {
        throw bb::Error(E_NOTIMPL, std::string("Code for ") + function.name + " is not readable");
    }
    if (std::memcmp(code, function.signature, function.signatureSize) != 0) {
        std::ostringstream out;
        out << "Signature mismatch for " << function.name << " at OART+0x" << std::hex
            << function.rva << "; refusing to call it";
        throw bb::Error(E_NOTIMPL, out.str());
    }
}

void* GuardedAddress(std::uintptr_t moduleBase, const GuardedFunction& function) {
    RequireSignature(moduleBase, function);
    return reinterpret_cast<void*>(moduleBase + function.rva);
}

std::uint32_t IntrusiveCount(const void* object) {
    if (!object) {
        return 0;
    }
    return LoadDword(object, kIntrusiveCountOffset);
}

void ReleaseIntrusive(void* object) noexcept {
    if (!object) {
        return;
    }
    auto vtable = *reinterpret_cast<void* const* const*>(object);
    using Release = void(__stdcall*)(void*);
    auto release = reinterpret_cast<Release>(*reinterpret_cast<void* const*>(
        reinterpret_cast<const std::uint8_t*>(vtable) + kIntrusiveReleaseSlot));
    release(object);
}

} // namespace bb::oart
