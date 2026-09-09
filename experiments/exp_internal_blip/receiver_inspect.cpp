/**
 * @file receiver_inspect.cpp
 * Guarded, read-only inspection of the OART objects behind a PowerPoint
 * FillFormat.
 *
 * This calls no private Office function. It only reads memory that a chain of
 * identity checks has already proved to be the expected object, and it fails
 * closed at the first check that does not match. Nothing observed here is
 * retained after the call returns.
 *
 * Why the offsets are believed to mean what they mean, on OART
 * 16.0.14334.20848 x64 (see docs/receiver_lookup.md for the full derivation):
 *
 *   - OART +0xAF60B8 is an IDispatch vtable: slot 0 compares IIDs, slot 1 is
 *     `++*(uint32*)(this+0x30)`, slot 2 is the matching decrement that destroys
 *     through the second base at this+8. Slot 17 is +0x8A13E0, the automation
 *     entry that performs Fill.UserPicture.
 *   - The object is 0x68 bytes. Its factory OART +0x23B220 allocates 0x68,
 *     installs that vtable, AddRefs a control block and stores it at +0x58.
 *   - The control block is the ordinary MSO shape: strong count at +0, weak
 *     count at +4, pointee at +0x10. OART +0x63EA0 is exactly that +0x10 load.
 *   - The pointee is the OART receiver, built by OART +0x223950 with vtable
 *     +0x9F6658, its slide container at +0x8 and an allocation sequence number
 *     at +0x20 taken from the global counter at OART +0xD40038.
 *
 * The vtable comparisons are the real guard. A build whose layout differs will
 * almost certainly fail them before any offset is used, and the exact-version
 * check refuses unknown builds outright.
 */

#include "../experiment_api.hpp"

#include <blipbridge/dispatch.hpp>

#include <cstdint>
#include <cstring>
#include <sstream>
#include <string>
#include <vector>

namespace {

/// Only this exact Office build has been observed and validated.
constexpr DWORD kSupportedVersionHigh = 0x00100000;   // 16.0
constexpr DWORD kSupportedVersionLow = (14334u << 16) | 20848u;
constexpr wchar_t kSupportedVersionText[] = L"16.0.14334.20848";

/// Observational RVAs and offsets. Never call targets; only compare and read.
///
/// PowerPoint's public FillFormat is a PPCORE object. Its vtable slot 17
/// (PPCORE +0x8249C0) is a forwarder that loads `this+0x08` and calls that
/// object's vtable at +0x88 - which is exactly the OART FillFormat slot holding
/// +0x8A13E0. So the inner object at +0x08 is the OART FillFormat.
constexpr uintptr_t kPublicFillFormatVtableRva = 0x1464478;  // ppcore.dll
constexpr size_t kPublicFillFormatSize = 0x10;
constexpr size_t kPublicFillFormatInnerOffset = 0x08;

constexpr uintptr_t kFillFormatVtableRva = 0xAF60B8;
constexpr uintptr_t kReceiverVtableRva = 0x9F6658;
constexpr size_t kFillFormatSize = 0x68;
constexpr size_t kFillFormatReferenceCountOffset = 0x30;
constexpr size_t kFillFormatTokenOffset = 0x58;
constexpr size_t kControlBlockSize = 0x18;
constexpr size_t kControlBlockStrongOffset = 0x00;
constexpr size_t kControlBlockPointeeOffset = 0x10;
constexpr size_t kReceiverInspectedSize = 0x28;
constexpr size_t kReceiverContainerOffset = 0x08;
constexpr size_t kReceiverSequenceOffset = 0x20;

/**
 * True when every page spanning [address, address+size) is committed and
 * readable. Used before each dereference so a wrong build cannot turn a
 * mismatched offset into an access violation.
 */
bool IsReadable(const void* address, size_t size) {
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

std::uintptr_t LoadPointer(const void* base, size_t offset) {
    std::uintptr_t value = 0;
    std::memcpy(&value, reinterpret_cast<const std::uint8_t*>(base) + offset, sizeof(value));
    return value;
}

std::uint32_t LoadDword(const void* base, size_t offset) {
    std::uint32_t value = 0;
    std::memcpy(&value, reinterpret_cast<const std::uint8_t*>(base) + offset, sizeof(value));
    return value;
}

/// Throws unless @p moduleName is loaded and is the build these offsets came from.
HMODULE RequireSupportedModule(const wchar_t* moduleName, const char* description) {
    HMODULE module = GetModuleHandleW(moduleName);
    if (!module) {
        throw bb::Error(E_NOTIMPL, std::string(description) + " is not loaded in this process");
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
                        "Receiver inspection supports only Office 16.0.14334.20848");
    }
    return module;
}

void AppendPointer(std::wostringstream& out, const wchar_t* label, std::uintptr_t value) {
    out << label << L"=0x" << std::hex << value << std::dec << L';';
}

/**
 * Describes an address as `module+0xRVA` when it lies in a loaded module.
 * Used only to make a failed identity check diagnosable on an unknown build;
 * the module handle is taken without incrementing its reference count and is
 * not retained.
 */
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

} // namespace

/**
 * Reports the OART receiver reachable from @p fill, or throws if any identity
 * check fails. The returned text is a semicolon-separated field list intended
 * for comparison across Shapes, not for parsing by production code.
 *
 * Ownership: every pointer reported is borrowed and belongs to Office. This
 * function takes no reference and stores nothing.
 */
std::wstring inspectFillReceiver(IDispatch* fill) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing FillFormat");
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Receiver inspection requires the PowerPoint host");
    }
    const auto oartBase = reinterpret_cast<std::uintptr_t>(
        RequireSupportedModule(L"oart.dll", "OART"));
    const auto ppcoreBase = reinterpret_cast<std::uintptr_t>(
        RequireSupportedModule(L"ppcore.dll", "PPCORE"));

    // Step 1: the argument must be PowerPoint's own FillFormat. A marshalling
    // proxy or a different build fails here and stops the walk.
    const void* publicFill = fill;
    if (!IsReadable(publicFill, kPublicFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "FillFormat storage is not readable");
    }
    const std::uintptr_t publicVtable = LoadPointer(publicFill, 0);
    if (publicVtable != ppcoreBase + kPublicFillFormatVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "FillFormat vtable is " + DescribeAddress(publicVtable) +
                            ", not the validated PPCORE+0x1464478 layout");
    }

    // Step 2: the OART FillFormat the public object forwards to.
    const auto handler = reinterpret_cast<const void*>(
        LoadPointer(publicFill, kPublicFillFormatInnerOffset));
    if (!IsReadable(handler, kFillFormatSize)) {
        throw bb::Error(E_NOTIMPL, "Inner FillFormat storage is not readable");
    }
    const std::uintptr_t handlerVtable = LoadPointer(handler, 0);
    if (handlerVtable != oartBase + kFillFormatVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "Inner FillFormat vtable is " + DescribeAddress(handlerVtable) +
                            ", not the validated OART+0xAF60B8 layout");
    }

    // Step 3: the control block held at +0x58.
    const std::uintptr_t token = LoadPointer(handler, kFillFormatTokenOffset);
    if (!IsReadable(reinterpret_cast<const void*>(token), kControlBlockSize)) {
        throw bb::Error(E_NOTIMPL, "Receiver control block is not readable");
    }
    const auto* tokenBytes = reinterpret_cast<const void*>(token);
    const std::uint32_t strongCount = LoadDword(tokenBytes, kControlBlockStrongOffset);
    const std::uintptr_t receiver = LoadPointer(tokenBytes, kControlBlockPointeeOffset);

    // Step 4: the receiver, confirmed by its own vtable before any field read.
    if (!IsReadable(reinterpret_cast<const void*>(receiver), kReceiverInspectedSize)) {
        throw bb::Error(E_NOTIMPL, "Receiver storage is not readable");
    }
    const auto* receiverBytes = reinterpret_cast<const void*>(receiver);
    const std::uintptr_t receiverVtable = LoadPointer(receiverBytes, 0);
    if (receiverVtable != oartBase + kReceiverVtableRva) {
        throw bb::Error(E_NOTIMPL,
                        "Receiver vtable is " + DescribeAddress(receiverVtable) +
                            ", not the validated OART+0x9F6658 layout");
    }

    std::wostringstream out;
    out << L"build=" << kSupportedVersionText << L';';
    AppendPointer(out, L"oart", oartBase);
    AppendPointer(out, L"publicFill", reinterpret_cast<std::uintptr_t>(publicFill));
    AppendPointer(out, L"handler", reinterpret_cast<std::uintptr_t>(handler));
    out << L"handlerRefs=" << LoadDword(handler, kFillFormatReferenceCountOffset) << L';';
    AppendPointer(out, L"token", token);
    out << L"tokenStrong=" << strongCount << L';';
    AppendPointer(out, L"receiver", receiver);
    AppendPointer(out, L"container", LoadPointer(receiverBytes, kReceiverContainerOffset));
    out << L"sequence=" << LoadPointer(receiverBytes, kReceiverSequenceOffset) << L';';
    return out.str();
}
