/**
 * @file unsupported_backend.cpp
 * The backend used where no accelerated implementation exists yet.
 *
 * Compiled instead of an accelerated backend wherever one has not been validated.
 * Two cases reach it today, and each gets its own reason rather than a shared
 * vague one:
 *
 *  - **Windows 32-bit.** The library builds, but the 32-bit PowerPoint backend
 *    has not been reverse-engineered or validated. The x64 implementation is not
 *    portable to it by recompilation: it depends on the x64 calling convention,
 *    on per-build module RVAs, and on object layouts that 32-bit Office does not
 *    share. Its sources are excluded from an x86 build entirely, so there is no
 *    possibility of a half-working backend reaching a document.
 *  - **Every other platform**, macOS being the case that matters.
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
