#pragma once
#include <windows.h>

namespace bb {
/** Invalid compressed image or image exceeding the experimental input limits. */
inline constexpr HRESULT BB_E_INVALID_IMAGE =
    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202);
/** Unknown, released, or foreign Engine handle. Handles are never recycled. */
inline constexpr HRESULT BB_E_TEXTURE_NOT_FOUND =
    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x206);
inline constexpr char kTextureNotFoundMessage[] = "Texture handle not found";
inline constexpr char kMemoryBackendUnavailableMessage[] =
    "No validated memory-image backend on this Office build. "
    "RegisterTextureShape is the explicit fallback.";
} // namespace bb
