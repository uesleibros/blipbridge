/**
 * @file unsupported_backend.cpp
 * The backend used where no accelerated implementation exists yet.
 *
 * Compiled where no backend can fill a Shape at all. **No configuration builds
 * it today**: Windows requires one of the two real backends, and Windows is the
 * only platform CMake will configure for.
 *
 * It is kept for macOS, which is the case that will need it, and it is the shape
 * that answer should take - see docs/macos.md. Windows 32-bit used to reach it
 * and no longer does: it has the portable backend, which actually works.
 *
 * It deliberately does nothing except answer honestly: no capabilities, and a
 * specific reason on every texture call. The library still loads, still reports
 * its version, and still explains itself, which is far more useful to a caller
 * than a missing export or a silent no-op.
 */

#include "backend.hpp"

namespace bb {
namespace {

class UnsupportedBackend final : public Backend {
  public:
    const char* Name() const noexcept override {
#if defined(_WIN32) && !defined(_WIN64)
        return "windows-x86-unvalidated";
#else
        return "unsupported-platform";
#endif
    }

    BackendResult Probe() noexcept override {
        return Refuse();
    }

    BackendCapabilities Capabilities() const noexcept override {
        return BackendCapabilities{}; // everything false, including the fallback
    }

    BackendResult
    LoadTexture(const std::uint8_t*, std::size_t, std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        return Refuse();
    }

    BackendResult LoadTexturePixels(const std::uint8_t*,
                                    std::uint32_t,
                                    std::uint32_t,
                                    std::int32_t,
                                    std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        return Refuse();
    }

    BackendResult ApplyTexture(void*, std::uint64_t) noexcept override {
        return Refuse();
    }

    BackendResult ApplyTextureIfChanged(void*, std::uint64_t, bool* skipped) noexcept override {
        if (skipped) {
            *skipped = false;
        }
        return Refuse();
    }

    BackendResult ApplyTextureRange(void*, std::uint64_t, std::uint32_t* applied) noexcept override {
        if (applied) {
            *applied = 0;
        }
        return Refuse();
    }

    BackendResult ReleaseTexture(std::uint64_t) noexcept override {
        return Refuse();
    }

    void ClearTextures() noexcept override {}

    BackendResult ApplyPicture(void*, const std::uint16_t*) noexcept override {
        return Refuse();
    }

    BackendResult LoadTexturePixelsScaled(const std::uint8_t*,
                                          std::uint32_t,
                                          std::uint32_t,
                                          std::int32_t,
                                          std::uint32_t,
                                          std::uint32_t,
                                          std::uint32_t,
                                          std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        return Refuse();
    }

    BackendResult InvalidateShape(void*) noexcept override {
        return Refuse();
    }

    void ClearPictureCache() noexcept override {}

    void PictureCacheStats(std::size_t* textures,
                           std::size_t* shapes,
                           std::uint64_t* skipped) const noexcept override {
        if (textures) {
            *textures = 0;
        }
        if (shapes) {
            *shapes = 0;
        }
        if (skipped) {
            *skipped = 0;
        }
    }

    std::size_t TextureCount() const noexcept override {
        return 0;
    }

  private:
    static BackendResult Refuse() {
        return BackendResult::Failure(BackendStatus::UnsupportedHost, kReason);
    }

#if defined(_WIN32) && !defined(_WIN64)
    static constexpr const char* kReason =
        "BlipBridge's accelerated backend is not available on 32-bit Windows yet. "
        "The 32-bit PowerPoint internals have not been validated, and the 64-bit "
        "implementation is not portable to them by recompilation, so it is not "
        "built here at all. Use 64-bit PowerPoint with the x64 package, or track "
        "docs/windows_x86.md";
#else
    static constexpr const char* kReason =
        "BlipBridge has no accelerated backend on this platform. The Windows "
        "implementation depends on PowerPoint internals that do not transfer; "
        "see docs/macos.md";
#endif
};

} // namespace

std::unique_ptr<Backend> CreateBackend() {
    return std::make_unique<UnsupportedBackend>();
}

} // namespace bb
