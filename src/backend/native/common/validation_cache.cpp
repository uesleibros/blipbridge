/**
 * @file validation_cache.cpp
 * The verdict store. See the header for what is cached and what never is.
 */

#include "validation_cache.hpp"

#include <sstream>

namespace bb::native {
namespace {

constexpr wchar_t kKeyPath[] = L"Software\\BlipBridge\\NativeValidation";

HKEY OpenKey(bool forWriting) noexcept {
    HKEY key = nullptr;
    if (forWriting) {
        if (RegCreateKeyExW(HKEY_CURRENT_USER,
                            kKeyPath,
                            0,
                            nullptr,
                            REG_OPTION_NON_VOLATILE,
                            KEY_WRITE,
                            nullptr,
                            &key,
                            nullptr) != ERROR_SUCCESS) {
            return nullptr;
        }
    } else if (RegOpenKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, KEY_READ, &key) != ERROR_SUCCESS) {
        return nullptr;
    }
    return key;
}

} // namespace

std::wstring ValidationKey(const NativeProfile& profile) {
    /*
     * Every module, in the order the profile lists them, with its full identity.
     * The resolver version and the architecture are in the key rather than
     * checked separately, so that a change to either simply misses instead of
     * matching an entry recorded under different rules.
     */
    std::wostringstream out;
    out << L"v" << kValidationCacheVersion << L'|' << Name(profile.architecture()) << L'|';
    for (const ModuleIdentity& module : profile.modules) {
        out << module.name << L'@' << module.VersionText() << L'@' << std::hex << module.timestamp
            << L'@' << module.imageSize << std::dec << L'@' << module.BuildSignatureText() << L';';
    }
    return out.str();
}

CachedVerdict LookUpValidation(const NativeProfile& profile) noexcept {
    if (!profile.valid()) {
        return CachedVerdict::Unknown;
    }

    /*
     * A profile whose modules are not all strongly identified is never looked
     * up, and so can never be skipped. Caching a verdict against a weak key
     * would let an Office rebuild that happened to keep its timestamp and size
     * inherit a proof it never earned.
     */
    for (const ModuleIdentity& module : profile.modules) {
        if (!module.strong()) {
            return CachedVerdict::Unknown;
        }
    }

    HKEY key = OpenKey(false);
    if (!key) {
        return CachedVerdict::Unknown;
    }
    const std::wstring name = ValidationKey(profile);
    DWORD value = 0;
    DWORD size = sizeof(value);
    DWORD type = 0;
    const LSTATUS status = RegQueryValueExW(
        key, name.c_str(), nullptr, &type, reinterpret_cast<LPBYTE>(&value), &size);
    RegCloseKey(key);

    if (status != ERROR_SUCCESS || type != REG_DWORD || size != sizeof(value)) {
        return CachedVerdict::Unknown;
    }
    // Only the exact recorded value counts as a pass. Anything else - a
    // truncated write, a hand-edited entry, a future encoding - reads as
    // "unknown", which costs a self-test rather than skipping one.
    return value == 1 ? CachedVerdict::Passed : CachedVerdict::Unknown;
}

void RecordValidationPassed(const NativeProfile& profile) noexcept {
    if (!profile.valid()) {
        return;
    }
    for (const ModuleIdentity& module : profile.modules) {
        if (!module.strong()) {
            return;
        }
    }

    HKEY key = OpenKey(true);
    if (!key) {
        // Nowhere to write is not an error worth surfacing: the only cost is
        // that the self-test runs again next time, which is the safe direction.
        return;
    }
    const std::wstring name = ValidationKey(profile);
    const DWORD value = 1;
    RegSetValueExW(
        key, name.c_str(), 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(value));
    RegCloseKey(key);
}

void ForgetValidationCache() noexcept {
    RegDeleteKeyExW(HKEY_CURRENT_USER, kKeyPath, KEY_WOW64_32KEY, 0);
    RegDeleteKeyExW(HKEY_CURRENT_USER, kKeyPath, KEY_WOW64_64KEY, 0);
    RegDeleteKeyW(HKEY_CURRENT_USER, kKeyPath);
}

} // namespace bb::native
