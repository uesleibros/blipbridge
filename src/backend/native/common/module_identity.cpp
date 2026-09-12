/**
 * @file module_identity.cpp
 * Reading a module's identity. See the header for why each field is in it.
 */

#include "module_identity.hpp"

#include "memory_safety.hpp"
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

} // namespace

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
