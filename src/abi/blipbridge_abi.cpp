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
#include "../image/pipeline.hpp"
#include "../image/store.hpp"
#include "../image/warp.hpp"

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

BB_API BB_Result BB_CALL BB_ApplyTextureIfChanged(void* shape,
                                                  BB_Handle texture,
                                                  int32_t* skipped) {
    try {
        if (skipped) {
            *skipped = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shape) {
            return Fail(BB_E_INVALID_ARG, "BB_ApplyTextureIfChanged needs a Shape pointer");
        }
        if (texture == 0) {
            return Fail(BB_E_INVALID_HANDLE, "Texture handle 0 is never valid");
        }
        bool avoided = false;
        const bb::BackendResult result =
            Library::Instance().Ensure().ApplyTextureIfChanged(shape, texture, &avoided);
        if (skipped && result.ok()) {
            *skipped = avoided ? 1 : 0;
        }
        return Translate(result);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyTextureIfChanged");
    }
}


/* ---------------------------------------------------------------------------
 * Images
 *
 * These are ordinary image work with no Office in them, so they are implemented
 * here over src/image/ rather than behind the backend seam - there is nothing
 * host-specific for a backend to decide. They still require BB_Init, because a
 * library that answers some calls before initialisation and not others is a
 * contract nobody can remember.
 * ------------------------------------------------------------------------- */

namespace {

/// Translates a caller's request struct, rejecting values we do not implement.
bool ReadRequest(const BB_ImageRequest* request, bb::image::PrepareRequest& out) {
    if (!request) {
        return true; // a null request changes nothing, which is a valid ask
    }
    if (!bb::image::IsSupportedTransform(request->transform) ||
        !bb::image::IsSupportedFilter(request->filter)) {
        return false;
    }
    out.cropX = request->cropX;
    out.cropY = request->cropY;
    out.cropWidth = request->cropWidth;
    out.cropHeight = request->cropHeight;
    out.transform = static_cast<bb::image::Transform>(request->transform);
    out.targetWidth = request->targetWidth;
    out.targetHeight = request->targetHeight;
    out.filter = static_cast<bb::image::ScaleFilter>(request->filter);
    return true;
}

/// Maps a prepare failure onto the ABI's error space, keeping the specific reason.
BB_Result FailPrepare(const bb::image::PrepareResult& result, const char* what) {
    std::string message = what;
    message += ": ";
    switch (result.status) {
    case bb::image::PrepareStatus::DecodeFailed:
        message += bb::image::DescribeDecodeStatus(result.decode);
        return Fail(result.decode == bb::image::DecodeStatus::NoData ? BB_E_FILE_NOT_FOUND
                                                                    : BB_E_DECODE_FAILED,
                    message);
    case bb::image::PrepareStatus::ResampleFailed:
        message += bb::image::DescribeStatus(result.resample);
        return Fail(BB_E_INVALID_ARG, message);
    default:
        message += bb::image::DescribePrepareStatus(result.status);
        return Fail(BB_E_INVALID_ARG, message);
    }
}

/// Registers a prepared image and hands back its handle.
BB_Result StoreImage(bb::image::PrepareResult& prepared, BB_Image* out) {
    auto owned = std::make_shared<bb::image::DecodedImage>(std::move(prepared.image));
    const std::uint64_t handle = bb::image::AddImage(std::move(owned));
    if (handle == 0) {
        return Fail(BB_E_INTERNAL, "The decoded image could not be registered");
    }
    *out = handle;
    g_lastError.clear();
    return BB_OK;
}

/// The image behind a handle, or a refusal that says which kind of handle it was.
BB_Result ResolveImage(BB_Image image, bb::image::ImageRef& out) {
    if (image == 0) {
        return Fail(BB_E_INVALID_HANDLE, "Image handle 0 is never valid");
    }
    if (!bb::image::IsImageHandle(image)) {
        return Fail(BB_E_INVALID_HANDLE,
                    "That is not an image handle. Image handles come from BB_LoadImage; "
                    "a texture handle belongs to BB_ApplyTexture.");
    }
    out = bb::image::FindImage(image);
    if (!out) {
        return Fail(BB_E_INVALID_HANDLE, "That image handle is not valid (it was released)");
    }
    return BB_OK;
}

} // namespace

BB_API BB_Result BB_CALL BB_LoadImage(const uint8_t* bytes,
                                      uint32_t length,
                                      const BB_ImageRequest* request,
                                      BB_Image* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!bytes || length == 0 || !out) {
            return Fail(BB_E_INVALID_ARG, "BB_LoadImage needs bytes, a length and an output");
        }
        bb::image::PrepareRequest prepare;
        if (!ReadRequest(request, prepare)) {
            return Fail(BB_E_INVALID_ARG, "That transform or filter is not one this build has");
        }
        bb::image::PrepareResult prepared = bb::image::PrepareEncoded(bytes, length, prepare);
        if (!prepared.ok()) {
            return FailPrepare(prepared, "BB_LoadImage");
        }
        return StoreImage(prepared, out);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadImage");
    }
}

BB_API BB_Result BB_CALL BB_LoadImageFromFile(const uint16_t* path,
                                              const BB_ImageRequest* request,
                                              BB_Image* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!path || !*path || !out) {
            return Fail(BB_E_INVALID_ARG, "BB_LoadImageFromFile needs a path and an output");
        }
        bb::image::PrepareRequest prepare;
        if (!ReadRequest(request, prepare)) {
            return Fail(BB_E_INVALID_ARG, "That transform or filter is not one this build has");
        }
        bb::image::PrepareResult prepared =
            bb::image::PrepareFile(reinterpret_cast<const wchar_t*>(path), prepare);
        if (!prepared.ok()) {
            return FailPrepare(prepared, "BB_LoadImageFromFile");
        }
        return StoreImage(prepared, out);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadImageFromFile");
    }
}

BB_API BB_Result BB_CALL BB_LoadImagePixels(const uint8_t* pixels,
                                            uint32_t width,
                                            uint32_t height,
                                            int32_t stride,
                                            const BB_ImageRequest* request,
                                            BB_Image* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!pixels || width == 0 || height == 0 || !out) {
            return Fail(BB_E_INVALID_ARG,
                        "BB_LoadImagePixels needs pixels, dimensions and an output");
        }
        if (stride < static_cast<int32_t>(width) * 4) {
            return Fail(BB_E_INVALID_ARG, "Stride must be at least width*4 bytes for BGRA32");
        }
        bb::image::PrepareRequest prepare;
        if (!ReadRequest(request, prepare)) {
            return Fail(BB_E_INVALID_ARG, "That transform or filter is not one this build has");
        }
        bb::image::PrepareResult prepared =
            bb::image::PreparePixels(pixels, width, height, stride, prepare);
        if (!prepared.ok()) {
            return FailPrepare(prepared, "BB_LoadImagePixels");
        }
        return StoreImage(prepared, out);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadImagePixels");
    }
}

BB_API BB_Result BB_CALL BB_GetImageSize(BB_Image image, uint32_t* width, uint32_t* height) {
    try {
        if (width) {
            *width = 0;
        }
        if (height) {
            *height = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        bb::image::ImageRef found;
        if (const BB_Result resolved = ResolveImage(image, found); resolved != BB_OK) {
            return resolved;
        }
        if (width) {
            *width = found->width;
        }
        if (height) {
            *height = found->height;
        }
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_GetImageSize");
    }
}

BB_API BB_Result BB_CALL BB_ReleaseImage(BB_Image image) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (image == 0) {
            return Fail(BB_E_INVALID_HANDLE, "Image handle 0 is never valid");
        }
        if (!bb::image::ReleaseImage(image)) {
            return Fail(BB_E_INVALID_HANDLE,
                        "That image handle is not valid (it was released, or it is a texture)");
        }
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ReleaseImage");
    }
}

BB_API BB_Result BB_CALL BB_ClearImages(void) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        bb::image::ClearImages();
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ClearImages");
    }
}

BB_API uint32_t BB_CALL BB_GetImageCount(void) {
    try {
        return bb::image::ImageCount();
    } catch (...) {
        return 0;
    }
}

BB_API BB_Result BB_CALL BB_CreateTextureFromImage(BB_Image image, BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!out) {
            return Fail(BB_E_INVALID_ARG, "BB_CreateTextureFromImage needs an output handle");
        }
        bb::image::ImageRef found;
        if (const BB_Result resolved = ResolveImage(image, found); resolved != BB_OK) {
            return resolved;
        }
        std::uint64_t handle = 0;
        const bb::BackendResult result = Library::Instance().Ensure().LoadTexturePixels(
            found->pixels.data(), found->width, found->height, found->stride(), &handle);
        if (!result.ok()) {
            return Translate(result);
        }
        *out = handle;
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_CreateTextureFromImage");
    }
}

namespace {

/// Warps an image onto four points, leaving the result in @p warped.
BB_Result WarpInto(BB_Image image,
                   const BB_PointF* points,
                   uint32_t filter,
                   bb::image::WarpResult& warped) {
    if (!points) {
        return Fail(BB_E_INVALID_ARG, "A quad needs four points");
    }
    if (!bb::image::IsSupportedFilter(filter)) {
        return Fail(BB_E_INVALID_ARG, "That filter is not one this build has");
    }
    bb::image::ImageRef found;
    if (const BB_Result resolved = ResolveImage(image, found); resolved != BB_OK) {
        return resolved;
    }

    bb::image::PointF quad[4];
    for (int index = 0; index < 4; ++index) {
        quad[index].x = points[index].x;
        quad[index].y = points[index].y;
    }

    warped = bb::image::WarpQuad(found->pixels.data(),
                                 found->width,
                                 found->height,
                                 found->stride(),
                                 quad,
                                 0,
                                 0,
                                 static_cast<bb::image::ScaleFilter>(filter));
    if (!warped.ok()) {
        return Fail(BB_E_INVALID_ARG,
                    std::string("BB_WarpImageQuad: ") +
                        bb::image::DescribeWarpStatus(warped.status));
    }
    return BB_OK;
}

} // namespace

BB_API BB_Result BB_CALL BB_WarpImageQuad(BB_Image image,
                                          const BB_PointF* points,
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
            return Fail(BB_E_INVALID_ARG, "BB_WarpImageQuad needs an output handle");
        }
        bb::image::WarpResult warped;
        if (const BB_Result status = WarpInto(image, points, filter, warped); status != BB_OK) {
            return status;
        }
        std::uint64_t handle = 0;
        const bb::BackendResult result =
            Library::Instance().Ensure().LoadTexturePixels(warped.image.pixels.data(),
                                                           warped.image.width,
                                                           warped.image.height,
                                                           warped.image.stride(),
                                                           &handle);
        if (!result.ok()) {
            return Translate(result);
        }
        *out = handle;
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_WarpImageQuad");
    }
}

BB_API BB_Result BB_CALL BB_ApplyImageQuad(void* shape,
                                           BB_Image image,
                                           const BB_PointF* points,
                                           uint32_t filter) {
    try {
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shape) {
            return Fail(BB_E_INVALID_ARG, "BB_ApplyImageQuad needs a Shape pointer");
        }

        // The warp comes first so that a bad quad is refused before a texture
        // exists to leak, and before the Shape is touched at all.
        BB_Handle texture = 0;
        if (const BB_Result warped = BB_WarpImageQuad(image, points, filter, &texture);
            warped != BB_OK) {
            return warped;
        }

        const BB_Result applied = BB_ApplyTexture(shape, texture);
        // The texture was made for this one apply, so it goes whether or not the
        // apply worked. Keeping it would leak an image per failed call.
        const std::string reason = applied == BB_OK ? std::string() : g_lastError;
        BB_ReleaseTexture(texture);
        if (applied != BB_OK) {
            g_lastError = reason;
            return applied;
        }
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyImageQuad");
    }
}

namespace {

/// Prepares an encoded image and registers it as a texture, never as an image.
BB_Result LoadTexturePrepared(bb::image::PrepareResult& prepared,
                              const char* what,
                              BB_Handle* out) {
    if (!prepared.ok()) {
        return FailPrepare(prepared, what);
    }
    std::uint64_t handle = 0;
    const bb::BackendResult result =
        Library::Instance().Ensure().LoadTexturePixels(prepared.image.pixels.data(),
                                                       prepared.image.width,
                                                       prepared.image.height,
                                                       prepared.image.stride(),
                                                       &handle);
    if (!result.ok()) {
        return Translate(result);
    }
    *out = handle;
    g_lastError.clear();
    return BB_OK;
}

} // namespace

BB_API BB_Result BB_CALL BB_LoadTextureEx(const uint8_t* bytes,
                                          uint32_t length,
                                          const BB_ImageRequest* request,
                                          BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!bytes || length == 0 || !out) {
            return Fail(BB_E_INVALID_ARG, "BB_LoadTextureEx needs bytes, a length and an output");
        }
        bb::image::PrepareRequest prepare;
        if (!ReadRequest(request, prepare)) {
            return Fail(BB_E_INVALID_ARG, "That transform or filter is not one this build has");
        }
        bb::image::PrepareResult prepared = bb::image::PrepareEncoded(bytes, length, prepare);
        return LoadTexturePrepared(prepared, "BB_LoadTextureEx", out);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadTextureEx");
    }
}

BB_API BB_Result BB_CALL BB_LoadTextureFromFileEx(const uint16_t* path,
                                                  const BB_ImageRequest* request,
                                                  BB_Handle* out) {
    try {
        if (out) {
            *out = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!path || !*path || !out) {
            return Fail(BB_E_INVALID_ARG, "BB_LoadTextureFromFileEx needs a path and an output");
        }
        bb::image::PrepareRequest prepare;
        if (!ReadRequest(request, prepare)) {
            return Fail(BB_E_INVALID_ARG, "That transform or filter is not one this build has");
        }
        bb::image::PrepareResult prepared =
            bb::image::PrepareFile(reinterpret_cast<const wchar_t*>(path), prepare);
        return LoadTexturePrepared(prepared, "BB_LoadTextureFromFileEx", out);
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_LoadTextureFromFileEx");
    }
}

BB_API BB_Result BB_CALL BB_ApplyTextureRange(void* shapeRange,
                                              BB_Handle texture,
                                              uint32_t* applied) {
    try {
        if (applied) {
            *applied = 0;
        }
        if (const BB_Result ready = RequireReadyThread(); ready != BB_OK) {
            return ready;
        }
        if (!shapeRange) {
            return Fail(BB_E_INVALID_ARG, "BB_ApplyTextureRange needs a ShapeRange pointer");
        }
        if (texture == 0) {
            return Fail(BB_E_INVALID_HANDLE, "Texture handle 0 is never valid");
        }
        std::uint32_t filled = 0;
        const bb::BackendResult result =
            Library::Instance().Ensure().ApplyTextureRange(shapeRange, texture, &filled);
        if (!result.ok()) {
            return Translate(result);
        }
        if (applied) {
            *applied = filled;
        }
        g_lastError.clear();
        return BB_OK;
    } catch (...) {
        return Fail(BB_E_INTERNAL, "Unknown failure during BB_ApplyTextureRange");
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
        bits |= capabilities.rangeApply ? BB_CAP_RANGE_APPLY : 0u;
        bits |= capabilities.imagePipeline ? BB_CAP_IMAGE_PIPELINE : 0u;
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
