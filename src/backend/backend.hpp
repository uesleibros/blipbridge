#pragma once
/**
 * @file backend.hpp
 * The seam between the public C ABI and whatever can actually fill a Shape.
 *
 * Everything above this line is platform-independent: the C ABI knows about
 * handles, byte buffers and error codes, and nothing about Office. Everything
 * below it is host-specific, and on Windows that means reverse-engineered
 * PowerPoint internals which must not leak upwards.
 *
 * There is exactly one backend per process, chosen at build time by
 * `CreateBackend`. A host with no accelerated implementation gets a backend
 * that answers honestly rather than one that pretends: it reports no
 * capabilities and refuses every texture operation with a documented reason.
 *
 * Ownership and threading are the ABI's contract, restated here because the
 * backend has to honour them:
 *
 *  - `shape` is borrowed for the duration of a call and is never stored.
 *  - `bytes` is copied before `LoadTexture` returns.
 *  - Handles are owned by the backend until released; they never recycle.
 *  - All calls arrive on the thread that initialised the library.
 */

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>

namespace bb {

/// Why a backend refused, in terms the C ABI can map to an error code.
enum class BackendStatus {
    Ok,
    UnsupportedHost,  ///< not running inside the host application
    UnsupportedBuild, ///< host present, but its version is not validated
    InvalidArgument,
    InvalidHandle,
    InvalidShape,
    UnsupportedShapeClass, ///< the Shape's class has no picture-fill path at all
    FileNotFound,          ///< the image path could not be read
    FallbackFailed,        ///< Fill.UserPicture itself refused
    DecodeFailed,
    ApplyFailed,
    OutOfMemory,
    Internal,
};

/// Result of a backend call: a status plus a message worth showing a human.
struct BackendResult {
    BackendStatus status = BackendStatus::Ok;
    std::string message;

    static BackendResult Success() {
        return BackendResult{};
    }

    static BackendResult Failure(BackendStatus status, std::string message) {
        return BackendResult{status, std::move(message)};
    }

    bool ok() const {
        return status == BackendStatus::Ok;
    }
};

/// Capability bits, mirroring the BB_CAP_* values in the public header.
struct BackendCapabilities {
    bool applyPicture = false; ///< BB_ApplyPicture and its caches are usable
    bool rangeApply = false;   ///< BB_ApplyTextureRange fills a whole ShapeRange
    bool imagePipeline = false; ///< BB_Image, crop/transform/scale and quad warp
    bool scaledPixels = false; ///< LoadTexturePixelsScaled is implemented
    bool nativeBackend = false;
    bool memoryImage = false;
    bool cachedTexture = false;
    bool batchApply = false;
    bool pickUpFallback = false;
    bool rawPixels = false; ///< LoadTexturePixels is implemented
};

/**
 * A host-specific implementation of the texture operations.
 *
 * Implementations must not throw: every failure comes back as a BackendResult.
 * The C ABI depends on that, because exceptions may not cross it.
 */
class Backend {
  public:
    virtual ~Backend() = default;

    /// Human-readable identifier, for example "windows-office-native".
    virtual const char* Name() const noexcept = 0;

    /// Probes the host. Called by BB_Init and safe to call repeatedly.
    virtual BackendResult Probe() noexcept = 0;

    virtual BackendCapabilities Capabilities() const noexcept = 0;

    /// Decodes @p bytes; on success writes a non-zero handle to @p out.
    virtual BackendResult
    LoadTexture(const std::uint8_t* bytes, std::size_t length, std::uint64_t* out) noexcept = 0;

    /**
     * Builds a texture from raw 32-bit BGRA pixels, skipping image decoding.
     * @p stride is bytes per row and may exceed width*4.
     */
    virtual BackendResult LoadTexturePixels(const std::uint8_t* pixels,
                                            std::uint32_t width,
                                            std::uint32_t height,
                                            std::int32_t stride,
                                            std::uint64_t* out) noexcept = 0;

    /// Fills the Shape behind @p shape, borrowed for this call only.
    virtual BackendResult ApplyTexture(void* shape, std::uint64_t texture) noexcept = 0;

    /**
     * Fills the Shape behind @p shape unless it already carries that image.
     *
     * @p skipped receives whether the Office edit was avoided. A skip leaves the
     * document untouched: no undo entry, no invalidation, no modified flag.
     */
    virtual BackendResult
    ApplyTextureIfChanged(void* shape, std::uint64_t texture, bool* skipped) noexcept = 0;

    /**
     * Fills every Shape in one ShapeRange with @p texture, through Office's own
     * range receiver rather than one apply per Shape.
     *
     * @p shapeRange is a PowerPoint `ShapeRange`, borrowed for the call.
     * @p applied receives how many member Shapes were filled.
     *
     * All or nothing: every member is classified first, and one member without a
     * validated native path refuses the whole range with nothing applied.
     */
    virtual BackendResult ApplyTextureRange(void* shapeRange,
                                            std::uint64_t texture,
                                            std::uint32_t* applied) noexcept = 0;

    virtual BackendResult ReleaseTexture(std::uint64_t texture) noexcept = 0;
    virtual void ClearTextures() noexcept = 0;

    /**
     * One-call picture fill: native where the Shape's class has a validated
     * path, ordinary host API where it does not, a specific refusal otherwise.
     * @p path is UTF-16 because that is what the host's file APIs take.
     */
    virtual BackendResult ApplyPicture(void* shape, const std::uint16_t* path) noexcept = 0;
    /**
     * Resamples BGRA32 pixels to a chosen size, then loads them as a texture.
     *
     * @p filter is a BB_SCALE_* value. Separate from LoadTexturePixels so that
     * neither entry point has ambiguous behaviour.
     */
    virtual BackendResult LoadTexturePixelsScaled(const std::uint8_t* pixels,
                                                  std::uint32_t width,
                                                  std::uint32_t height,
                                                  std::int32_t stride,
                                                  std::uint32_t targetWidth,
                                                  std::uint32_t targetHeight,
                                                  std::uint32_t filter,
                                                  std::uint64_t* out) noexcept = 0;

    /// Forgets what was last applied to one Shape.
    virtual BackendResult InvalidateShape(void* shape) noexcept = 0;
    /// Releases every path-keyed texture and forgets every Shape.
    virtual void ClearPictureCache() noexcept = 0;
    /// Cached textures, remembered Shapes, and applies skipped since Init.
    virtual void PictureCacheStats(std::size_t* textures,
                                   std::size_t* shapes,
                                   std::uint64_t* skipped) const noexcept = 0;
    virtual std::size_t TextureCount() const noexcept = 0;
};

/**
 * Builds the backend for this platform.
 *
 * Never returns null: where no accelerated implementation exists, this returns
 * one that reports no capabilities and explains why. That keeps
 * BB_GetCapabilities and BB_GetLastError meaningful everywhere.
 */
std::unique_ptr<Backend> CreateBackend();

} // namespace bb
