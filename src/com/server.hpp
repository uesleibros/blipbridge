#pragma once
#include <windows.h>
#include <atomic>

namespace bb {
inline constexpr CLSID kEngineClsid = {
    0x2e2e2731, 0xc523, 0x486b, {0x89, 0xcb, 0x2a, 0x89, 0x48, 0x4f, 0x1e, 0x32}};

/** Process-wide COM server lifetime. Owns no Office objects or module handles. */
struct ServerLifetime {
    std::atomic<long> objects{0};
    std::atomic<long> locks{0};
};
ServerLifetime& GetServerLifetime() noexcept;
/** Returns one owned Engine interface reference; exceptions never escape. */
HRESULT CreateEngine(REFIID interfaceId, void** result) noexcept;
} // namespace bb
