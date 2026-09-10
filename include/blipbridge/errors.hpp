#pragma once
#include <windows.h>

namespace bb {
/** Invalid compressed image or image exceeding the experimental input limits. */
inline constexpr HRESULT BB_E_INVALID_IMAGE = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x202);
/** Unknown, released, or foreign Engine handle. Handles are never recycled. */
inline constexpr HRESULT BB_E_TEXTURE_NOT_FOUND = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x206);
/** The Shape's class has no picture-fill path at all - not native, not fallback. */
inline constexpr HRESULT BB_E_SHAPE_CLASS_UNSUPPORTED =
    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x207);
/** An image path that cannot be read: missing, a directory, or unreadable. */
inline constexpr HRESULT BB_E_IMAGE_FILE_MISSING =
    MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x208);
/** Office's own Fill.UserPicture refused, on a class that had nothing else left. */
inline constexpr HRESULT BB_E_FALLBACK_REFUSED = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, 0x209);
inline constexpr char kTextureNotFoundMessage[] = "Texture handle not found";
inline constexpr char kMemoryBackendUnavailableMessage[] =
    "No validated memory-image backend on this Office build. "
    "RegisterTextureShape is the explicit fallback.";
} // namespace bb
