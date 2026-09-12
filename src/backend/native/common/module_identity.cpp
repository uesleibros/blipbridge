/**
 * @file module_identity.cpp
 * Reading a module's identity. See the header for why each field is in it.
 */

#include "module_identity.hpp"

#include "memory_safety.hpp"
#include <cstring>
#include <iomanip>
#include <sstream>
#include <vector>

namespace bb::native {
namespace {

/// The NT headers of a mapped image, or null when the image is not well formed.
const IMAGE_NT_HEADERS* NtHeaders(HMODULE module) noexcept {
    const auto base = reinterpret_cast<const std::uint8_t*>(module);
    if (!base || !IsReadable(base, sizeof(IMAGE_DOS_HEADER))) {
        return nullptr;
    }
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew <= 0) {
        return nullptr;
    }
    const auto* headers = reinterpret_cast<const IMAGE_NT_HEADERS*>(base + dos->e_lfanew);
    if (!IsReadable(headers, sizeof(IMAGE_NT_HEADERS))) {
        return nullptr;
    }
    if (headers->Signature != IMAGE_NT_SIGNATURE) {
        return nullptr;
    }
    return headers;
}

/**
 * The four version numbers, from the module file's version resource.
 *
 * Read from disk rather than from the mapped image: the resource is in the file,
 * and GetFileVersionInfo is the documented way to it. A module with no version
 * resource leaves the numbers at zero, which is a fact about it rather than a
 * failure - the timestamp and image size still identify it.
 */
void ReadVersion(const std::wstring& path, ModuleIdentity& identity) noexcept {
    if (path.empty()) {
        return;
    }
    DWORD ignored = 0;
    const DWORD size = GetFileVersionInfoSizeW(path.c_str(), &ignored);
    if (size == 0) {
        return;
    }
    std::vector<std::uint8_t> buffer(size);
    if (!GetFileVersionInfoW(path.c_str(), 0, size, buffer.data())) {
        return;
    }
    VS_FIXEDFILEINFO* fixed = nullptr;
    UINT length = 0;
    if (!VerQueryValueW(buffer.data(), L"\\", reinterpret_cast<void**>(&fixed), &length) ||
        !fixed || length < sizeof(VS_FIXEDFILEINFO)) {
        return;
    }
    identity.major = static_cast<std::uint16_t>(HIWORD(fixed->dwFileVersionMS));
    identity.minor = static_cast<std::uint16_t>(LOWORD(fixed->dwFileVersionMS));
    identity.build = static_cast<std::uint16_t>(HIWORD(fixed->dwFileVersionLS));
    identity.revision = static_cast<std::uint16_t>(LOWORD(fixed->dwFileVersionLS));
}

/// The CodeView record a PDB-linked image carries in its debug directory.
struct CodeViewRsds {
    std::uint32_t signature; // 'RSDS'
    std::uint8_t guid[16];
    std::uint32_t age;
    // A NUL-terminated PDB path follows. Deliberately not read: the path is a
    // build-machine detail that says nothing about the image's contents.
};

constexpr std::uint32_t kRsds = 0x53445352; // 'RSDS' little-endian

/**
 * Reads the build GUID and age out of @p module's debug directory.
 *
 * Read from the **mapped image**, not from the file. The debug directory sits in
 * a read-only section and carries no relocations, so the mapped bytes are the
 * file's bytes - which is what makes this stable across loads. That is not true
 * of code on x86, where relocation rewrites absolute addresses throughout .text
 * and a hash of the mapped code would differ from one process to the next.
 */
void ReadBuildSignature(HMODULE module,
                        const IMAGE_NT_HEADERS* headers,
                        ModuleIdentity& identity) noexcept {
    const auto base = reinterpret_cast<const std::uint8_t*>(module);
    const IMAGE_DATA_DIRECTORY& directory =
        headers->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DEBUG];
    if (directory.VirtualAddress == 0 || directory.Size < sizeof(IMAGE_DEBUG_DIRECTORY)) {
        return;
    }

    const auto* entries =
        reinterpret_cast<const IMAGE_DEBUG_DIRECTORY*>(base + directory.VirtualAddress);
    const std::size_t count = directory.Size / sizeof(IMAGE_DEBUG_DIRECTORY);
    if (!IsReadable(entries, directory.Size)) {
        return;
    }

    for (std::size_t index = 0; index < count; ++index) {
        const IMAGE_DEBUG_DIRECTORY& entry = entries[index];
        if (entry.Type != IMAGE_DEBUG_TYPE_CODEVIEW || entry.AddressOfRawData == 0) {
            continue;
        }
        if (entry.SizeOfData < sizeof(CodeViewRsds)) {
            continue;
        }
        const auto* record = reinterpret_cast<const CodeViewRsds*>(base + entry.AddressOfRawData);
        if (!IsReadable(record, sizeof(CodeViewRsds)) || record->signature != kRsds) {
            continue;
        }
        std::memcpy(identity.buildGuid, record->guid, sizeof(identity.buildGuid));
        identity.buildAge = record->age;
        identity.hasBuildSignature = true;
        return;
    }
}

} // namespace

std::wstring ModuleIdentity::BuildSignatureText() const {
    if (!hasBuildSignature) {
        return std::wstring();
    }
    // Printed in the GUID's own field order, so it can be compared by eye against
    // a symbol server or a debugger's module list.
    const auto data1 = static_cast<std::uint32_t>(buildGuid[0]) |
                       (static_cast<std::uint32_t>(buildGuid[1]) << 8) |
                       (static_cast<std::uint32_t>(buildGuid[2]) << 16) |
                       (static_cast<std::uint32_t>(buildGuid[3]) << 24);
    const auto data2 = static_cast<std::uint16_t>(buildGuid[4] | (buildGuid[5] << 8));
    const auto data3 = static_cast<std::uint16_t>(buildGuid[6] | (buildGuid[7] << 8));

    std::wostringstream out;
    out << std::hex << std::uppercase << std::setfill(L'0');
    out << L'{' << std::setw(8) << data1 << L'-' << std::setw(4) << data2 << L'-' << std::setw(4)
        << data3 << L'-';
    for (int index = 8; index < 10; ++index) {
        out << std::setw(2) << static_cast<unsigned>(buildGuid[index]);
    }
    out << L'-';
    for (int index = 10; index < 16; ++index) {
        out << std::setw(2) << static_cast<unsigned>(buildGuid[index]);
    }
    out << L"}+" << std::dec << buildAge;
    return out.str();
}

std::wstring ModuleIdentity::VersionText() const {
    if (major == 0 && minor == 0 && build == 0 && revision == 0) {
        return std::wstring();
    }
    std::wostringstream out;
    out << major << L'.' << minor << L'.' << build << L'.' << revision;
    return out.str();
}

std::wstring ModuleIdentity::Describe() const {
    std::wostringstream out;
    out << (name.empty() ? L"<unnamed>" : name) << L' ' << Name(architecture);
    const std::wstring version = VersionText();
    if (!version.empty()) {
        out << L' ' << version;
    }
    out << L" ts=0x" << std::hex << timestamp << L" size=0x" << imageSize << std::dec;
    const std::wstring signature = BuildSignatureText();
    if (!signature.empty()) {
        out << L' ' << signature;
    } else {
        // Said out loud, because an identity with no build signature cannot
        // authorise a native profile and a reader needs to know that is why.
        out << L" (no build signature)";
    }
    return out.str();
}

ModuleIdentity IdentifyModule(HMODULE module) noexcept {
    ModuleIdentity identity;
    if (!module) {
        return identity;
    }

    const IMAGE_NT_HEADERS* headers = NtHeaders(module);
    if (!headers) {
        // Not a well-formed mapped image. Left invalid rather than partially
        // filled: a caller must not be able to match on a name alone.
        return identity;
    }

    identity.architecture = ArchitectureOf(module);
    identity.timestamp = headers->FileHeader.TimeDateStamp;
    identity.imageSize = headers->OptionalHeader.SizeOfImage;
    ReadBuildSignature(module, headers, identity);

    wchar_t path[MAX_PATH] = {};
    const DWORD length = GetModuleFileNameW(module, path, MAX_PATH);
    if (length > 0 && length < MAX_PATH) {
        const std::wstring full(path, length);
        const std::size_t slash = full.find_last_of(L"\\/");
        identity.name = slash == std::wstring::npos ? full : full.substr(slash + 1);
        ReadVersion(full, identity);
    }

    return identity;
}

ModuleIdentity IdentifyLoadedModule(const wchar_t* name) noexcept {
    if (!name) {
        return ModuleIdentity();
    }
    // GetModuleHandleW, never LoadLibrary: identifying a module must not have
    // the side effect of bringing one into the process.
    return IdentifyModule(GetModuleHandleW(name));
}

} // namespace bb::native
