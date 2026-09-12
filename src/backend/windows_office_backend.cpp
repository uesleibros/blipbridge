/**
 * @file windows_office_backend.cpp
 * The Windows PowerPoint backend.
 *
 * This is the only file that knows the public API is sitting on top of
 * reverse-engineered Office internals. It adapts the validated native texture
 * store to the platform-independent `Backend` interface, and it converts that
 * store's exceptions into `BackendResult` values so nothing can escape into the
 * C ABI.
 *
 * The safety model is unchanged and lives one level down, in
 * `src/backend/windows_office`: Office build checks, structural PPCORE
 * wrapper validation, OART and GFX vtable checks, per-function signature bytes,
 * and receiver re-resolution on every apply with no cached receiver pointer. A
 * host this backend does not recognise fails closed, and the failure surfaces
 * here as UnsupportedHost or UnsupportedBuild rather than as a crash.
 */

#include "../image/resample.hpp"
#include "backend.hpp"
#include "backend_guard.hpp"
#include "windows_office/native_texture.hpp"
#include "windows_office/picture_cache.hpp"
#include "windows_office/range_texture.hpp"
#include "windows_office/shape_policy.hpp"
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstring>
#include <new>
// MinGW requires Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

namespace bb {
namespace {

/**
 * The Windows implementation.
 *
 * It holds no state of its own: the texture store lives in the native module so
 * that the COM Automation surface and this C ABI address the same textures and
 * cannot disagree about what is loaded.
 */
class WindowsOfficeBackend final : public Backend {
  public:
    const char* Name() const noexcept override {
        return "windows-office-native";
    }

    BackendResult Probe() noexcept override {
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return BackendResult::Failure(
                BackendStatus::UnsupportedHost,
                "BlipBridge's accelerated backend runs inside PowerPoint for Windows; "
                "this process is not PowerPoint");
        }
        if (!nativeTextureBackendAvailable()) {
            return BackendResult::Failure(
                BackendStatus::UnsupportedBuild,
                "This Office build has not been validated. The accelerated backend "
                "depends on internal layouts that are checked per build, and refuses "
                "to run against an unrecognised one");
        }
        return BackendResult::Success();
    }

    BackendCapabilities Capabilities() const noexcept override {
        BackendCapabilities capabilities;
        // Every capability here needs a live PowerPoint, including the donor
        // fallback, which drives PickUp/Apply through Automation. Reporting any
        // of them in a process that is not PowerPoint would be a lie a caller
        // could act on.
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return capabilities;
        }
        capabilities.pickUpFallback = true;
        // The picture path works wherever PowerPoint does: on a build the native
        // backend cannot serve, it still dispatches every Shape to Office's own
        // Fill.UserPicture, which is a real capability rather than a stub.
        capabilities.applyPicture = true;
        const bool available = nativeTextureBackendAvailable();
        capabilities.nativeBackend = available;
        capabilities.memoryImage = available;
        capabilities.cachedTexture = available;
        capabilities.batchApply = available;
        // The range apply is the native path applied to a range receiver, so it
        // exists exactly when the native path does.
        capabilities.rangeApply = available;
        // The image work itself is portable and would run anywhere, but the end
        // of the pipeline is a texture, so the capability means what a caller
        // cares about: can I decode, crop, warp and get something Office takes.
        capabilities.imagePipeline = available;
        capabilities.rawPixels = available;
        capabilities.scaledPixels = available;
        return capabilities;
    }

    BackendResult LoadTexture(const std::uint8_t* bytes,
                              std::size_t length,
                              std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        if (!bytes || length == 0 || !out) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Image bytes and an output handle are required");
        }
        if (length > 0xFFFFFFFFull) {
            return BackendResult::Failure(BackendStatus::InvalidArgument, "Image is too large");
        }
        return Guarded([&] {
            // The store copies the bytes, so this array only has to outlive the call.
            SAFEARRAYBOUND bound{static_cast<ULONG>(length), 0};
            SAFEARRAY* array = SafeArrayCreate(VT_UI1, 1, &bound);
            if (!array) {
                throw std::bad_alloc();
            }
            struct Destroy {
                SAFEARRAY* value;
                ~Destroy() {
                    SafeArrayDestroy(value);
                }
            } destroy{array};

            void* raw = nullptr;
            check(SafeArrayAccessData(array, &raw), "SafeArrayAccessData");
            std::memcpy(raw, bytes, length);
            SafeArrayUnaccessData(array);
            *out = static_cast<std::uint64_t>(nativeTextureLoad(array));
        });
    }

    BackendResult LoadTexturePixels(const std::uint8_t* pixels,
                                    std::uint32_t width,
                                    std::uint32_t height,
                                    std::int32_t stride,
                                    std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        if (!pixels || width == 0 || height == 0 || !out) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Pixels, dimensions and an output handle are required");
        }
        // A stride smaller than one row of BGRA would read past every row.
        if (stride < static_cast<std::int32_t>(width) * 4) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Stride must be at least width*4 bytes for BGRA32");
        }
        return Guarded([&] {
            *out =
                static_cast<std::uint64_t>(nativeTextureLoadPixels(pixels, width, height, stride));
        });
    }

    BackendResult ApplyTexture(void* shape, std::uint64_t texture) noexcept override {
        return Guarded([&] {
            long shapeType = 0;
            IDispatch* dispatch = RequireFillableShape(shape, &shapeType);
            // Fill is fetched per call; the receiver behind it is resolved inside
            // and never cached, because a deleted Shape still passes every check.
            nativeTextureApplyToShape(dispatch, static_cast<long>(texture), shapeType);
        });
    }

    BackendResult
    ApplyTextureIfChanged(void* shape, std::uint64_t texture, bool* skipped) noexcept override {
        return Guarded([&] {
            // Same gate, same order: an ineligible Shape is refused before the
            // skip cache is consulted, so this cannot become a way to reach the
            // private backend with a Shape the ordinary apply would reject.
            long shapeType = 0;
            IDispatch* dispatch = RequireFillableShape(shape, &shapeType);
            nativeTextureApplyIfChanged(dispatch, static_cast<long>(texture), shapeType, skipped);
        });
    }

    BackendResult ApplyTextureRange(void* shapeRange,
                                    std::uint64_t texture,
                                    std::uint32_t* applied) noexcept override {
        return Guarded([&] {
            // A ShapeRange is not a Shape, so the Shape-class gate cannot be
            // asked about it here: the pointer is checked, and every *member* is
            // classified inside, before anything internal is touched.
            IDispatch* dispatch = RequireDispatchShape(shapeRange);
            ReleaseDispatch release{dispatch};

            const std::uint32_t filled =
                office::ApplyTextureToRange(dispatch, static_cast<long>(texture));
            if (applied) {
                *applied = filled;
            }
        });
    }

    BackendResult ReleaseTexture(std::uint64_t texture) noexcept override {
        return Guarded([&] { nativeTextureRelease(static_cast<long>(texture)); });
    }

    void ClearTextures() noexcept override {
        nativeTextureClear();
    }

    BackendResult LoadTexturePixelsScaled(const std::uint8_t* pixels,
                                          std::uint32_t width,
                                          std::uint32_t height,
                                          std::int32_t stride,
                                          std::uint32_t targetWidth,
                                          std::uint32_t targetHeight,
                                          std::uint32_t filter,
                                          std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        if (!out) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "An output handle is required");
        }
        // Every argument is checked by the resampler, which reports precisely
        // which one was wrong; repeating those checks here would only let the
        // two disagree.
        std::vector<std::uint8_t> scaled;
        const image::ResampleStatus status =
            image::Resample(pixels,
                            width,
                            height,
                            stride,
                            targetWidth,
                            targetHeight,
                            static_cast<image::ScaleFilter>(filter),
                            scaled);
        if (status != image::ResampleStatus::Ok) {
            const BackendStatus code = status == image::ResampleStatus::OutOfMemory
                                           ? BackendStatus::OutOfMemory
                                           : BackendStatus::InvalidArgument;
            return BackendResult::Failure(code, image::DescribeStatus(status));
        }
        return Guarded([&] {
            // The resampled buffer is tightly packed, so its stride is exactly
            // one row of BGRA.
            *out = static_cast<std::uint64_t>(nativeTextureLoadPixels(
                scaled.data(), targetWidth, targetHeight, static_cast<long>(targetWidth) * 4));
        });
    }

    BackendResult ApplyPicture(void* shape, const std::uint16_t* path) noexcept override {
        if (!path || !*path) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "An image path is required");
        }
        return Guarded([&] {
            // The Shape is validated as a pointer here; which *classes* are
            // acceptable is decided inside, because the picture path also has a
            // fallback for classes the native path refuses.
            IDispatch* dispatch = RequireDispatchShape(shape);
            ReleaseDispatch release{dispatch};
            office::ApplyPictureCached(dispatch, reinterpret_cast<const wchar_t*>(path));
        });
    }

    BackendResult InvalidateShape(void* shape) noexcept override {
        return Guarded([&] {
            IDispatch* dispatch = RequireDispatchShape(shape);
            ReleaseDispatch release{dispatch};
            office::InvalidateShapeCache(dispatch);
        });
    }

    void ClearPictureCache() noexcept override {
        office::ClearPictureCache();
    }

    void PictureCacheStats(std::size_t* textures,
                           std::size_t* shapes,
                           std::uint64_t* skipped) const noexcept override {
        const office::PictureCacheStats stats = office::GetPictureCacheStats();
        if (textures) {
            *textures = stats.textures;
        }
        if (shapes) {
            *shapes = stats.shapes;
        }
        if (skipped) {
            *skipped = stats.skipped;
        }
    }

    std::size_t TextureCount() const noexcept override {
        return static_cast<std::size_t>(nativeTextureCount());
    }
};

} // namespace

std::unique_ptr<Backend> CreateBackend() {
    return std::make_unique<WindowsOfficeBackend>();
}

} // namespace bb
