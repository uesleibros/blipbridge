/**
 * @file memory_safety.cpp
 * Guarded reads. See the header for why nothing here trusts a pointer.
 */

#include "memory_safety.hpp"

#include <cstring>

namespace bb::native {
namespace {

/// The protections under which a page cannot be read at all.
constexpr DWORD kUnreadable = PAGE_NOACCESS | PAGE_GUARD;

constexpr DWORD kExecutable =
    PAGE_EXECUTE | PAGE_EXECUTE_READ | PAGE_EXECUTE_READWRITE | PAGE_EXECUTE_WRITECOPY;

} // namespace

bool IsReadable(const void* address, std::size_t size) noexcept {
    if (!address || size == 0) {
        return false;
    }

    const auto* cursor = static_cast<const std::uint8_t*>(address);
    const std::uint8_t* end = cursor + size;
    if (end < cursor) {
        // The range wrapped, so the size is nonsense and no read from it is safe.
        return false;
    }

    /*
     * Walked region by region rather than page by page. A range can span several
     * VirtualQuery regions with different protections, and checking only the
     * first page would pass a buffer whose tail is on a guard page - which is
     * exactly the read that would fault.
     */
    while (cursor < end) {
        MEMORY_BASIC_INFORMATION information{};
        if (VirtualQuery(cursor, &information, sizeof(information)) != sizeof(information)) {
            return false;
        }
        if (information.State != MEM_COMMIT) {
            return false;
        }
        if (information.Protect == 0 || (information.Protect & kUnreadable) != 0) {
            return false;
        }
        const auto* regionEnd =
            static_cast<const std::uint8_t*>(information.BaseAddress) + information.RegionSize;
        if (regionEnd <= cursor) {
            // No forward progress: refuse rather than loop.
            return false;
        }
        cursor = regionEnd;
    }
    return true;
}

bool IsExecutable(const void* address) noexcept {
    if (!address) {
        return false;
    }
    MEMORY_BASIC_INFORMATION information{};
    if (VirtualQuery(address, &information, sizeof(information)) != sizeof(information)) {
        return false;
    }
    return information.State == MEM_COMMIT && (information.Protect & kExecutable) != 0;
}

std::uintptr_t ReadPointer(const void* base, std::size_t offset) noexcept {
    if (!base) {
        return 0;
    }
    const auto* at = static_cast<const std::uint8_t*>(base) + offset;
    std::uintptr_t value = 0;
    if (!IsReadable(at, sizeof(value))) {
        return 0;
    }
    std::memcpy(&value, at, sizeof(value));
    return value;
}

std::uint32_t ReadDword(const void* base, std::size_t offset) noexcept {
    if (!base) {
        return 0;
    }
    const auto* at = static_cast<const std::uint8_t*>(base) + offset;
    std::uint32_t value = 0;
    if (!IsReadable(at, sizeof(value))) {
        return 0;
    }
    std::memcpy(&value, at, sizeof(value));
    return value;
}

bool BytesMatch(const void* address, const std::uint8_t* expected, std::size_t size) noexcept {
    if (!address || !expected || size == 0) {
        return false;
    }
    if (!IsReadable(address, size)) {
        return false;
    }
    return std::memcmp(address, expected, size) == 0;
}

} // namespace bb::native
