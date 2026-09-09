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

#include <psapi.h>

#include <cstring>
#include <map>
#include <set>
#include <sstream>
#include <vector>

namespace bb::oart {
namespace {

/// The wrapper vtable observed for Shape.Fill on the tested build; diagnostics only.
constexpr std::uintptr_t kObservedFillFormatVtableRva = 0x1464478;   // ppcore.dll
constexpr std::size_t kPublicFillFormatSize = 0x10;
/// How many identity thunks a vtable must have before it counts as a wrapper.
/// The observed Shape.Fill table has 23; the sibling wrappers have 10 to 31.
constexpr unsigned kMinimumIdentityThunks = 6;
/// How far into a vtable to look. The observed tables delegate within this range.
constexpr unsigned kWrapperSlotsExamined = 48;

constexpr std::uintptr_t kFillFormatVtableRva = 0xAF60B8;          // oart.dll
constexpr std::size_t kFillFormatSize = 0x68;
constexpr std::size_t kFillFormatTokenOffset = 0x58;
constexpr std::size_t kFillFormatFlagOffset = 0x60;

constexpr std::size_t kControlBlockSize = 0x18;
constexpr std::size_t kControlBlockStrongOffset = 0x00;
constexpr std::size_t kControlBlockPointeeOffset = 0x10;

constexpr std::uintptr_t kReceiverVtableRva = 0x9F6658;            // oart.dll
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
bool DecodeIdentityThunk(const std::uint8_t* code, std::size_t& innerOffset,
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

    void RememberModule(HMODULE module) { validatedModules_.insert(module); }

    /// Looks up a previously validated vtable shape, keyed by its address.
    const DelegatingWrapper* FindWrapper(std::uintptr_t vtable) const {
        const auto found = wrappers_.find(vtable);
        return found == wrappers_.end() ? nullptr : &found->second;
    }

    void RememberWrapper(std::uintptr_t vtable, const DelegatingWrapper& wrapper) {
        wrappers_.emplace(vtable, wrapper);
    }

    /// Drops everything if any module of interest was unloaded or moved.
    void SynchroniseWith(HMODULE oart, HMODULE ppcore, HMODULE gfx) {
        if (oart == oart_ && ppcore == ppcore_ && gfx == gfx_) {
            return;
        }
        validatedModules_.clear();
        wrappers_.clear();
        oart_ = oart;
        ppcore_ = ppcore;
        gfx_ = gfx;
    }

private:
    ValidationCache() = default;

    std::set<HMODULE> validatedModules_;
    std::map<std::uintptr_t, DelegatingWrapper> wrappers_;
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

bool DescribeDelegatingWrapper(const void* object, std::uintptr_t moduleBase,
                               std::size_t moduleSize, DelegatingWrapper& wrapper) {
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
    if (!IsReadable(reinterpret_cast<const void*>(vtable),
                    kWrapperSlotsExamined * sizeof(void*))) {
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
        if (!DecodeIdentityThunk(reinterpret_cast<const std::uint8_t*>(entry), innerOffset,
                                 slotOffset)) {
            continue;
        }
        if (slotOffset != slot * sizeof(void*)) {
            continue;   // delegates, but not to the matching slot
        }
        if (haveOffset && innerOffset != wrapper.innerOffset) {
            return false;
        }
        wrapper.innerOffset = innerOffset;
        haveOffset = true;
        ++wrapper.identityThunks;
    }
    if (!haveOffset || wrapper.identityThunks < kMinimumIdentityThunks) {
        return false;   // not a wrapper; nothing worth remembering
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
        cursor = reinterpret_cast<const std::uint8_t*>(information.BaseAddress) +
                 information.RegionSize;
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
                            reinterpret_cast<LPCWSTR>(address), &owner) ||
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


HMODULE RequireSupportedModule(const wchar_t* moduleName, const char* description) {
    HMODULE module = GetModuleHandleW(moduleName);
    if (!module) {
        throw bb::Error(E_NOTIMPL,
                        std::string(description) +
                            " is not loaded; warm it with an ordinary picture fill first");
    }
    ValidationCache& cache = ValidationCache::Instance();
    cache.SynchroniseWith(GetModuleHandleW(L"oart.dll"), GetModuleHandleW(L"ppcore.dll"),
                          GetModuleHandleW(L"gfx.dll"));
    if (cache.IsModuleValidated(module)) {
        return module;   // same loaded image; its version cannot have changed
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

    FillTarget target;
    target.oartBase =
        reinterpret_cast<std::uintptr_t>(RequireSupportedModule(L"oart.dll", "OART"));
    target.ppcoreBase =
        reinterpret_cast<std::uintptr_t>(RequireSupportedModule(L"ppcore.dll", "PPCORE"));

    // Step 1: the argument must be a PPCORE delegating wrapper. This is a
    // structural test, not a vtable address: PPCORE has a family of these and
    // pinning one RVA rejects Shape.Fill's own siblings.
    target.publicFill = fill;
    if (!IsReadable(target.publicFill, kPublicFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "FillFormat storage is not readable");
    }
    MODULEINFO ppcoreInfo{};
    if (!GetModuleInformation(GetCurrentProcess(), GetModuleHandleW(L"ppcore.dll"),
                              &ppcoreInfo, sizeof(ppcoreInfo))) {
        throw bb::Error(E_NOTIMPL, "Cannot measure the PPCORE image");
    }
    if (!DescribeDelegatingWrapper(target.publicFill, target.ppcoreBase,
                                   ppcoreInfo.SizeOfImage, target.wrapper)) {
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
            << " delegates through this+0x" << std::hex << target.wrapper.innerOffset
            << " to " << DescribeAddress(handlerVtable) << ", not the validated OART+0x"
            << kFillFormatVtableRva
            << " FillFormat layout. Shape.Line and Shape.TextFrame look like this.";
        throw bb::Error(E_NOTIMPL, out.str());
    }
    target.handlerFlag = static_cast<std::uint8_t>(
        LoadDword(target.handler, kFillFormatFlagOffset) & 0xFF);

    // Step 3: the control block held at +0x58.
    target.token = reinterpret_cast<void*>(LoadPointer(target.handler, kFillFormatTokenOffset));
    if (!IsReadable(target.token, kControlBlockSize)) {
        throw bb::Error(E_NOTIMPL, "Receiver control block is not readable");
    }
    target.tokenStrong = LoadDword(target.token, kControlBlockStrongOffset);

    // Step 4: the receiver, confirmed by its own vtable before any field read.
    target.receiver = reinterpret_cast<void*>(
        LoadPointer(target.token, kControlBlockPointeeOffset));
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
        throw bb::Error(E_NOTIMPL,
                        std::string("Code for ") + function.name + " is not readable");
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
