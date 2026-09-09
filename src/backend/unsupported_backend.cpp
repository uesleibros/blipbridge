/**
 * @file unsupported_backend.cpp
 * The backend used where no accelerated implementation exists yet.
 *
 * Compiled instead of the Windows backend on any other platform - macOS being
 * the case that matters. It deliberately does nothing except answer honestly:
 * no capabilities, and a specific reason on every texture call. The library
 * still loads, still reports its version, and still explains itself, which is
 * far more useful to a caller than a missing export or a silent no-op.
 *
 * What a macOS backend would need is genuinely unknown territory, not a port.
 * The Windows implementation is built on PPCORE/OART/GFX internals, the
 * Windows x64 ABI, and per-build module RVAs and signature bytes; none of that
 * transfers. See docs/macos.md for what would have to be researched first.
 */

#include "backend.hpp"

namespace bb {
namespace {

class UnsupportedBackend final : public Backend {
public:
    const char* Name() const noexcept override { return "unsupported-platform"; }

    BackendResult Probe() noexcept override { return Refuse(); }

    BackendCapabilities Capabilities() const noexcept override {
        return BackendCapabilities{};   // everything false, including the fallback
    }

    BackendResult LoadTexture(const std::uint8_t*, std::size_t,
                              std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        return Refuse();
    }

    BackendResult LoadTexturePixels(const std::uint8_t*, std::uint32_t, std::uint32_t,
                                    std::int32_t, std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        return Refuse();
    }

    BackendResult ApplyTexture(void*, std::uint64_t) noexcept override { return Refuse(); }
    BackendResult ReleaseTexture(std::uint64_t) noexcept override { return Refuse(); }
    void ClearTextures() noexcept override {}

    BackendResult ApplyPicture(void*, const std::uint16_t*) noexcept override {
        return Refuse();
    }
    BackendResult InvalidateShape(void*) noexcept override { return Refuse(); }
    void ClearPictureCache() noexcept override {}
    void PictureCacheStats(std::size_t* textures, std::size_t* shapes,
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
    std::size_t TextureCount() const noexcept override { return 0; }

private:
    static BackendResult Refuse() {
        return BackendResult::Failure(
            BackendStatus::UnsupportedHost,
            "BlipBridge has no accelerated backend on this platform. The Windows "
            "implementation depends on PowerPoint internals that do not transfer; "
            "see docs/macos.md");
    }
};

} // namespace

std::unique_ptr<Backend> CreateBackend() {
    return std::make_unique<UnsupportedBackend>();
}

} // namespace bb
