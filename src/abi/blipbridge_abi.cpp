/**
 * @file blipbridge_abi.cpp
 * Implementation of the public C ABI.
 *
 * This layer owns three responsibilities and nothing else:
 *
 *  1. Nothing escapes. Every entry point is noexcept in effect: the backend
 *     returns results rather than throwing, and the outermost catch here exists
 *     for the impossible case, because an exception crossing a C ABI is
 *     undefined behaviour rather than a bug report.
 *  2. Thread affinity. BB_Init records its thread and every later call is
 *     checked against it, since the Office objects underneath use non-atomic
 *     reference counts.
 *  3. Error text. Failures are stored per thread so BB_GetLastError can describe
 *     the most recent one without any allocation crossing the boundary.
 *
 * It knows nothing about Office. Everything host-specific is behind
 * bb::Backend.
 */

#include "../backend/backend.hpp"

#include <blipbridge/blipbridge.h>
#include <cstring>
#include <memory>
#include <mutex>
#include <string>

#if defined(_WIN32)
#include <windows.h>
#endif

namespace {

// The header is the single source of the release number; repeating it here is
// how the two drifted apart before.
constexpr uint32_t kVersionMajor = BB_VERSION_MAJOR;
constexpr uint32_t kVersionMinor = BB_VERSION_MINOR;
constexpr uint32_t kVersionPatch = BB_VERSION_PATCH;

#if defined(_WIN32) && defined(_M_X64) || defined(_WIN32) && defined(__x86_64__)
constexpr const char* kBuildTag = "windows-x64";
#elif defined(_WIN32)
constexpr const char* kBuildTag = "windows";
#elif defined(__APPLE__)
constexpr const char* kBuildTag = "macos";
#else
constexpr const char* kBuildTag = "unknown";
#endif

/// Per-thread description of the most recent failure, for BB_GetLastError.
thread_local std::string g_lastError;

unsigned long CurrentThread() {
#if defined(_WIN32)
    return static_cast<unsigned long>(GetCurrentThreadId());
#else
    return 0;
#endif
}

/**
 * Library state.
 *
 * The backend is created once and kept: constructing it is cheap, but the
 * texture store it fronts must be a single instance for handles to mean
 * anything. `owningThread` is zero until BB_Init succeeds in claiming it.
 */
struct Library {
    std::unique_ptr<bb::Backend> backend;
    unsigned long owningThread = 0;

    static Library& Instance() {
        static Library library;
        return library;
    }

    bb::Backend& Ensure() {
        if (!backend) {
            backend = bb::CreateBackend();
        }
        return *backend;
    }
};

void SetError(std::string message) {
    try {
        g_lastError = std::move(message);
    } catch (...) {
        // A failure to record why something failed must not itself fail.
    }
}

BB_Result Fail(BB_Result code, std::string message) {
    SetError(std::move(message));
    return code;
}

/// Maps the backend vocabulary onto the public error codes.
BB_Result CodeFor(bb::BackendStatus status) {
    switch (status) {
    case bb::BackendStatus::Ok:
        return BB_OK;
    case bb::BackendStatus::UnsupportedHost:
        return BB_E_UNSUPPORTED_HOST;
    case bb::BackendStatus::UnsupportedBuild:
        return BB_E_UNSUPPORTED_BUILD;
    case bb::BackendStatus::InvalidArgument:
        return BB_E_INVALID_ARG;
    case bb::BackendStatus::InvalidHandle:
        return BB_E_INVALID_HANDLE;
    case bb::BackendStatus::InvalidShape:
        return BB_E_INVALID_SHAPE;
    case bb::BackendStatus::UnsupportedShapeClass:
        return BB_E_UNSUPPORTED_SHAPE;
    case bb::BackendStatus::FileNotFound:
        return BB_E_FILE_NOT_FOUND;
    case bb::BackendStatus::FallbackFailed:
        return BB_E_FALLBACK_FAILED;
    case bb::BackendStatus::DecodeFailed:
        return BB_E_DECODE_FAILED;
    case bb::BackendStatus::ApplyFailed:
        return BB_E_APPLY_FAILED;
    case bb::BackendStatus::OutOfMemory:
        return BB_E_OUT_OF_MEMORY;
    case bb::BackendStatus::Internal:
        break;
    }
    return BB_E_INTERNAL;
}

BB_Result Translate(const bb::BackendResult& result) {
    if (result.ok()) {
        g_lastError.clear();
        return BB_OK;
    }
    return Fail(CodeFor(result.status), result.message);
}

/// Every entry point that touches textures starts here.
BB_Result RequireReadyThread() {
    Library& library = Library::Instance();
    if (library.owningThread == 0) {
        return Fail(BB_E_NOT_INITIALIZED, "BB_Init has not been called");
    }
    if (library.owningThread != CurrentThread()) {
        return Fail(BB_E_WRONG_THREAD,
                    "BlipBridge is single-threaded: call it from the thread that "
                    "called BB_Init");
    }
    return BB_OK;
}

/// Copies @p text as UTF-8 and reports the size the caller would need.
uint32_t CopyOut(const std::string& text, char* buffer, uint32_t capacity) {
    const uint32_t required = static_cast<uint32_t>(text.size() + 1);
    if (!buffer || capacity == 0) {
        return required;
    }
    const uint32_t copied = (required <= capacity) ? text.size() : (capacity - 1);
    std::memcpy(buffer, text.data(), copied);
    buffer[copied] = '\0';
    return required;
}

} // namespace

extern "C" {

BB_API BB_Result BB_CALL BB_Init(void) {
    try {
        Library& library = Library::Instance();
        const unsigned long thread = CurrentThread();
        if (library.owningThread != 0 && library.owningThread != thread) {
            return Fail(BB_E_WRONG_THREAD, "BlipBridge is already initialised on another thread");
        }
        bb::Backend& backend = library.Ensure();
        const bb::BackendResult probe = backend.Probe();
        // The thread is claimed even when the backend is unavailable, so that
        // later calls report the real reason instead of "not initialised".
        library.owningThread = thread;
        return Translate(probe);
    } catch (const std::exception& error) {
        return Fail(BB_E_INTERNAL, error.what());
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_Init");
    }
}

BB_API BB_Result BB_CALL BB_Shutdown(void) {
    try {
        Library& library = Library::Instance();
        if (library.backend) {
            // Both, because the two are now independently owned: caller handles
            // and the picture cache each reference their images, and clearing
            // one deliberately leaves the other alone. Shutdown is the only
            // place that must let go of everything BlipBridge owns.
            library.backend->ClearTextures();
            library.backend->ClearPictureCache();
        }
        library.owningThread = 0;
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_Shutdown");
    }
}

BB_API BB_Result BB_CALL BB_LoadTexture(const uint8_t* bytes, uint32_t length, BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!bytes || length == 0 || !out) {
            return Fail(BB_E_INVALID_ARG,
                        "BB_LoadTexture needs a non-empty byte buffer and an output handle");
        }
        std::uint64_t handle = 0;
        const bb::BackendResult result =
            Library::Instance().Ensure().LoadTexture(bytes, length, &handle);
        if (!result.ok()) {
            return Translate(result);
        }
        *out = handle;
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadTexture");
    }
}

BB_API BB_Result BB_CALL BB_LoadTexturePixels(
    const uint8_t* pixels, uint32_t width, uint32_t height, int32_t stride, BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!pixels || width == 0 || height == 0 || !out) {
            return Fail(BB_E_INVALID_ARG,
                        "BB_LoadTexturePixels needs pixels, dimensions and an output handle");
        }
        std::uint64_t handle = 0;
        const bb::BackendResult result =
            Library::Instance().Ensure().LoadTexturePixels(pixels, width, height, stride, &handle);
        if (!result.ok()) {
            return Translate(result);
        }
        *out = handle;
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadTexturePixels");
    }
}

BB_API BB_Result BB_CALL BB_ApplyTexture(void* shape, BB_Handle texture) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shape) {
            return Fail(BB_E_INVALID_ARG, "BB_ApplyTexture needs a Shape pointer");
        }
        if (texture == 0) {
            return Fail(BB_E_INVALID_HANDLE, "Texture handle 0 is never valid");
        }
        return Translate(Library::Instance().Ensure().ApplyTexture(shape, texture));
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyTexture");
    }
}

BB_API BB_Result BB_CALL BB_ApplyTextureBatch(void* const* shapes,
                                              const BB_Handle* textures,
                                              uint32_t count,
                                              uint32_t* applied) {
    try {
        if (applied) {
            *applied = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shapes || !textures) {
            return Fail(BB_E_INVALID_ARG,
                        "BB_ApplyTextureBatch needs both a Shape array and a handle array");
        }
        if (count == 0) {
            g_lastError.clear();
            return BB_OK;
        }
        bb::Backend& backend = Library::Instance().Ensure();
        for (uint32_t index = 0; index < count; ++index) {
            if (!shapes[index] || textures[index] == 0) {
                return Fail(BB_E_INVALID_ARG, "A batch entry has a null Shape or handle 0");
            }
            const bb::BackendResult result = backend.ApplyTexture(shapes[index], textures[index]);
            if (!result.ok()) {
                // Shapes already filled stay filled; the caller learns how far
                // the batch got from `applied`.
                return Translate(result);
            }
            if (applied) {
                *applied = index + 1;
            }
        }
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyTextureBatch");
    }
}

BB_API BB_Result BB_CALL BB_ReleaseTexture(BB_Handle texture) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (texture == 0) {
            return Fail(BB_E_INVALID_HANDLE, "Texture handle 0 is never valid");
        }
        return Translate(Library::Instance().Ensure().ReleaseTexture(texture));
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ReleaseTexture");
    }
}

BB_API BB_Result BB_CALL BB_ClearTextures(void) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        Library::Instance().Ensure().ClearTextures();
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ClearTextures");
    }
}

BB_API uint32_t BB_CALL BB_GetTextureCount(void) {
    try {
        return static_cast<uint32_t>(Library::Instance().Ensure().TextureCount());
    } catch (...) {
        return 0;
    }
}

BB_API BB_Result BB_CALL BB_LoadTexturePixelsScaled(const uint8_t* pixels,
                                                    uint32_t width,
                                                    uint32_t height,
                                                    int32_t stride,
                                                    uint32_t targetWidth,
                                                    uint32_t targetHeight,
                                                    uint32_t filter,
                                                    BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!out) {
            return Fail(BB_E_INVALID_ARG, "An output handle is required");
        }
        // Sizes, stride and filter are validated inside, which is where the
        // specific reason for each rejection lives.
        return Translate(Library::Instance().Ensure().LoadTexturePixelsScaled(
            pixels, width, height, stride, targetWidth, targetHeight, filter, out));
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadTexturePixelsScaled");
    }
}

BB_API BB_Result BB_CALL BB_ApplyPicture(void* shape, const uint16_t* path) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shape) {
            return Fail(BB_E_INVALID_ARG, "A Shape pointer is required");
        }
        if (!path || !*path) {
            return Fail(BB_E_INVALID_ARG, "An image path is required");
        }
        return Translate(Library::Instance().Ensure().ApplyPicture(shape, path));
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyPicture");
    }
}

BB_API BB_Result BB_CALL BB_InvalidateShape(void* shape) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shape) {
            return Fail(BB_E_INVALID_ARG, "A Shape pointer is required");
        }
        return Translate(Library::Instance().Ensure().InvalidateShape(shape));
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_InvalidateShape");
    }
}

BB_API BB_Result BB_CALL BB_ClearPictureCache(void) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        Library::Instance().Ensure().ClearPictureCache();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ClearPictureCache");
    }
}

BB_API BB_Result BB_CALL BB_GetPictureCacheStats(uint32_t* textures,
                                                 uint32_t* shapes,
                                                 uint64_t* skipped) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        std::size_t textureCount = 0;
        std::size_t shapeCount = 0;
        std::uint64_t skippedCount = 0;
        Library::Instance().Ensure().PictureCacheStats(&textureCount, &shapeCount, &skippedCount);
        if (textures) {
            *textures = static_cast<uint32_t>(textureCount);
        }
        if (shapes) {
            *shapes = static_cast<uint32_t>(shapeCount);
        }
        if (skipped) {
            *skipped = skippedCount;
        }
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_GetPictureCacheStats");
    }
}

BB_API uint32_t BB_CALL BB_GetCapabilities(void) {
    try {
        const bb::BackendCapabilities capabilities = Library::Instance().Ensure().Capabilities();
        uint32_t bits = 0;
        bits |= capabilities.nativeBackend ? BB_CAP_NATIVE_BACKEND : 0u;
        bits |= capabilities.memoryImage ? BB_CAP_MEMORY_IMAGE : 0u;
        bits |= capabilities.cachedTexture ? BB_CAP_CACHED_TEXTURE : 0u;
        bits |= capabilities.batchApply ? BB_CAP_BATCH_APPLY : 0u;
        bits |= capabilities.pickUpFallback ? BB_CAP_PICKUP_FALLBACK : 0u;
        bits |= capabilities.rawPixels ? BB_CAP_RAW_PIXELS : 0u;
        bits |= capabilities.applyPicture ? BB_CAP_APPLY_PICTURE : 0u;
        bits |= capabilities.scaledPixels ? BB_CAP_SCALED_PIXELS : 0u;
        return bits;
    } catch (...) {
        return 0;
    }
}

BB_API uint32_t BB_CALL BB_GetLastError(char* buffer, uint32_t capacity) {
    try {
        return CopyOut(g_lastError, buffer, capacity);
    } catch (...) {
        if (buffer && capacity > 0) {
            buffer[0] = '\0';
        }
        return 1;
    }
}

BB_API uint32_t BB_CALL BB_GetAbiVersion(void) {
    return BB_ABI_VERSION;
}

BB_API uint32_t BB_CALL BB_GetVersion(void) {
    return (kVersionMajor << 16) | (kVersionMinor << 8) | kVersionPatch;
}

BB_API uint32_t BB_CALL BB_GetVersionString(char* buffer, uint32_t capacity) {
    try {
        std::string text = std::to_string(kVersionMajor) + "." + std::to_string(kVersionMinor) +
                           "." + std::to_string(kVersionPatch) + " (" + kBuildTag + ", " +
                           Library::Instance().Ensure().Name() + ")";
        return CopyOut(text, buffer, capacity);
    } catch (...) {
        return CopyOut(std::string("unknown"), buffer, capacity);
    }
}

} // extern "C"
