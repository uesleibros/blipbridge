/**
 * @file architecture.cpp
 * PE machine reading. See architecture.hpp for why it is read rather than assumed.
 */

#include "architecture.hpp"

#include "memory_safety.hpp"

namespace bb::native {

const char* Name(Architecture architecture) noexcept {
    switch (architecture) {
    case Architecture::X86:
        return "x86";
    case Architecture::X64:
        return "x64";
    case Architecture::Unknown:
        break;
    }
    return "unknown";
}

Architecture ArchitectureOf(HMODULE module) noexcept {
    if (!module) {
        return Architecture::Unknown;
    }

    /*
     * Walked with guarded reads rather than by trusting that a HMODULE points at
     * a well-formed image. The header is read from a live process, and every
     * step here is one an unmapped or malformed module could fail at.
     */
    const auto base = reinterpret_cast<const std::uint8_t*>(module);
    if (!IsReadable(base, sizeof(IMAGE_DOS_HEADER))) {
        return Architecture::Unknown;
    }
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew <= 0) {
        return Architecture::Unknown;
    }

    const std::uint8_t* headers = base + dos->e_lfanew;
    // Only the signature and the file header are needed, and reading no further
    // keeps this correct for both PE32 and PE32+ without branching on either.
    constexpr std::size_t kSignatureAndFileHeader = sizeof(DWORD) + sizeof(IMAGE_FILE_HEADER);
    if (!IsReadable(headers, kSignatureAndFileHeader)) {
        return Architecture::Unknown;
    }
    if (*reinterpret_cast<const DWORD*>(headers) != IMAGE_NT_SIGNATURE) {
        return Architecture::Unknown;
    }

    const auto* file = reinterpret_cast<const IMAGE_FILE_HEADER*>(headers + sizeof(DWORD));
    switch (file->Machine) {
    case IMAGE_FILE_MACHINE_I386:
        return Architecture::X86;
    case IMAGE_FILE_MACHINE_AMD64:
        return Architecture::X64;
    default:
        // ARM64 and anything else: recognised as "not one of ours" rather than
        // mapped onto the nearest thing, because no profile exists for it.
        return Architecture::Unknown;
    }
}

bool MatchesCurrent(HMODULE module) noexcept {
    const Architecture actual = ArchitectureOf(module);
    return actual != Architecture::Unknown && actual == Current();
}

} // namespace bb::native
