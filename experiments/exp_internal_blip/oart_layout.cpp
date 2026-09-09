/**
 * @file oart_layout.cpp
 * Implementation of the guarded Office-internal accessors.
 *
 * The offsets below are observational on Office 16.0.14334.20848 x64. Their
 * derivation is in docs/receiver_lookup.md:
 *
 *  - PPCORE +0x8249C0 is the public FillFormat's vtable slot 17 and is a bare
 *    forwarder that loads `this+0x08` and calls that object's vtable at +0x88,
 *    which is where the OART FillFormat vtable holds the UserPicture entry. So
 *    the inner object at +0x08 is the OART FillFormat.
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
#include <sstream>
#include <vector>

namespace bb::oart {
namespace {

constexpr std::uintptr_t kPublicFillFormatVtableRva = 0x1464478;   // ppcore.dll
constexpr std::size_t kPublicFillFormatSize = 0x10;
constexpr std::size_t kPublicFillFormatInnerOffset = 0x08;

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

    // Step 1: the argument must be PowerPoint's own FillFormat. A marshalling
    // proxy or a different build fails here and stops the walk.
    target.publicFill = fill;
    if (!IsReadable(target.publicFill, kPublicFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "FillFormat storage is not readable");
    }
    const std::uintptr_t publicVtable = LoadPointer(target.publicFill, 0);
    if (publicVtable != target.ppcoreBase + kPublicFillFormatVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "FillFormat vtable is " + DescribeAddress(publicVtable) +
                            ", not the validated PPCORE+0x1464478 layout");
    }

    // Step 2: the OART FillFormat the public object forwards to.
    target.handler = reinterpret_cast<void*>(
        LoadPointer(target.publicFill, kPublicFillFormatInnerOffset));
    if (!IsReadable(target.handler, kFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "Inner FillFormat storage is not readable");
    }
    const std::uintptr_t handlerVtable = LoadPointer(target.handler, 0);
    if (handlerVtable != target.oartBase + kFillFormatVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "Inner FillFormat vtable is " + DescribeAddress(handlerVtable) +
                            ", not the validated OART+0xAF60B8 layout");
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
