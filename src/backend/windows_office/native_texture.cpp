/**
 * @file native_texture.cpp
 * Reusable texture handles over GFX cached images.
 *
 * A texture is one decoded image that any number of Shapes can be filled with.
 * `LoadTexture` decodes once; `ApplyTexture` then costs only record construction
 * and the receiver call.
 *
 * ## Ownership, which is the whole point of this file
 *
 * A cached image is a **GFX-level** object. `GEL::ICachedImage::Create` takes an
 * `IStream` and nothing else - no presentation, no slide, no Shape - so the
 * image does not belong to any document. A texture handle therefore owns exactly
 * one intrusive reference to it, and that reference is independent of whatever
 * documents happen to be open.
 *
 * What follows from that:
 *
 *  - A handle stays valid across `Presentation.Close`. The document releases its
 *    own references; ours is untouched.
 *  - A handle must not outlive the GFX module. Everything is released when the
 *    Engine disconnects or is destroyed, which happens while Office is still
 *    running. Leaving a reference outstanding at process teardown would be a
 *    leak at best.
 *  - Applying needs a live Shape, but holding a texture does not.
 *
 * Handles never recycle, so a stale handle is reported rather than silently
 * resolving to a different texture - the same rule the donor fallback follows.
 *
 * Threading: STA only, like everything else that touches these Office objects.
 * The store records its owning thread and refuses calls from any other.
 *
 * This is research-grade. `MemoryImageToFill`, `CachedTextureApply` and
 * `InternalBackend` stay false until the lifetime matrix in
 * tools/test_native_texture.ps1 has been run and the reference accounting is
 * fully attributed. See docs/native_texture.md.
 */

#include "native_texture.hpp"

#include "shape_policy.hpp"

#include "native_apply.hpp"
#include "oart_layout.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <map>
#include <memory>
#include <sstream>

namespace {

/**
 * Native texture handles start well above the donor fallback's, so one handle
 * space serves both backends and `ApplyTexture` can tell them apart without a
 * flag. The donor store hands out 1, 2, 3...; nothing will reach this base.
 */
constexpr long kNativeHandleBase = 0x1000000;

/// Same exported symbol the creator uses; see native_apply.cpp.
constexpr char kCreateFromStreamSymbol[] =
    "?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@"
    "AEAV?$TCntPtr@UIImage@GEL@@@4@PEAUIStream@@W4IStreamCopyInstruction@12@"
    "PEBVMD4UID@4@_N@Z";

/**
 * One decoded image, owning one reference to each of the GFX objects the
 * creator returned.
 */
class NativeTexture {
public:
    NativeTexture(void* cached, void* image, std::size_t byteCount)
        : cached_(cached), image_(image), byteCount_(byteCount) {}
    NativeTexture(const NativeTexture&) = delete;
    NativeTexture& operator=(const NativeTexture&) = delete;
    ~NativeTexture() {
        // Order mirrors creation: the image is the companion output, the cached
        // image is the one Office holds while a fill is in place.
        bb::oart::ReleaseIntrusive(image_);
        bb::oart::ReleaseIntrusive(cached_);
    }

    void* cached() const { return cached_; }
    std::size_t byteCount() const { return byteCount_; }
    unsigned long applyCount() const { return applyCount_; }
    void countApply() { ++applyCount_; }

private:
    void* cached_ = nullptr;
    void* image_ = nullptr;
    std::size_t byteCount_ = 0;
    unsigned long applyCount_ = 0;
};

/**
 * Process-wide texture store.
 *
 * The Engine is a single-threaded-apartment object and Office hosts one of it,
 * so a singleton is the honest model; it records the thread that created the
 * first texture and refuses every other thread. `Clear` is called from the
 * Engine's disconnect and destructor, which is what bounds the lifetime.
 */
class TextureStore {
public:
    static TextureStore& Instance() {
        static TextureStore store;
        return store;
    }

    long Add(void* cached, void* image, std::size_t byteCount) {
        RequireOwningThread(true);
        const long handle = nextHandle_++;
        textures_.emplace(handle, std::make_unique<NativeTexture>(cached, image, byteCount));
        ++creations_;
        return handle;
    }

    NativeTexture& Get(long handle) {
        RequireOwningThread(false);
        const auto found = textures_.find(handle);
        if (found == textures_.end()) {
            std::ostringstream out;
            out << "Texture handle " << handle << " is not valid";
            if (handle > 0 && handle < nextHandle_) {
                out << " (it was released; handles are never recycled)";
            }
            throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, out.str());
        }
        return *found->second;
    }

    void Release(long handle) {
        RequireOwningThread(false);
        if (textures_.erase(handle) == 0) {
            std::ostringstream out;
            out << "Texture handle " << handle << " is not valid";
            throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, out.str());
        }
    }

    /// Releases every texture. Safe to call when the store was never used.
    void Clear() noexcept {
        // Deliberately not thread-checked: this also runs from teardown paths.
        textures_.clear();
    }

    std::size_t Count() const { return textures_.size(); }
    long NextHandle() const { return nextHandle_; }

    /// Total decodes since the process started. A reuse claim is only credible
    /// if this stays far below the number of applies.
    unsigned long Creations() const { return creations_; }
    void CountCreation() { ++creations_; }

private:
    TextureStore() = default;

    void RequireOwningThread(bool mayClaim) {
        const DWORD current = GetCurrentThreadId();
        if (owningThread_ == 0) {
            if (!mayClaim) {
                throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, "No textures have been loaded");
            }
            owningThread_ = current;
            return;
        }
        if (owningThread_ != current) {
            throw bb::Error(RPC_E_WRONG_THREAD,
                            "Textures belong to the apartment that created them");
        }
    }

    std::map<long, std::unique_ptr<NativeTexture>> textures_;
    unsigned long creations_ = 0;
    long nextHandle_ = kNativeHandleBase;
    DWORD owningThread_ = 0;
};

} // namespace

bool nativeTextureBackendAvailable() noexcept {
    // Everything the backend needs, checked without touching a document: the
    // host, all three modules at the validated version, the structural checks
    // those imply, and every private entry point's signature bytes. Any failure
    // means the backend is unavailable here, which is the honest answer for a
    // machine running a different Office build.
    try {
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return false;
        }
        const HMODULE oart = bb::oart::RequireSupportedModule(L"oart.dll", "OART");
        bb::oart::RequireSupportedModule(L"ppcore.dll", "PPCORE");
        const HMODULE gfx = bb::oart::RequireSupportedModule(L"gfx.dll", "GFX");
        if (!GetProcAddress(gfx, kCreateFromStreamSymbol)) {
            return false;
        }
        bb::oart::ResolveApplyFunctions(reinterpret_cast<std::uintptr_t>(oart));
        return true;
    } catch (...) {
        return false;
    }
}

long nativeTextureLoad(SAFEARRAY* bytes) {
    LONG lower = 0;
    LONG upper = 0;
    if (!bytes || SafeArrayGetDim(bytes) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional Byte array");
    }
    bb::check(SafeArrayGetLBound(bytes, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(bytes, 1, &upper), "SafeArrayGetUBound");

    // Decoding is the expensive half, and it happens exactly once per texture.
    const bb::oart::CreatedImage created = bb::oart::CreateCachedImageFromBytes(bytes);
    try {
        return TextureStore::Instance().Add(created.cached, created.image,
                                            static_cast<std::size_t>(upper - lower) + 1);
    } catch (...) {
        bb::oart::ReleaseIntrusive(created.image);
        bb::oart::ReleaseIntrusive(created.cached);
        throw;
    }
}

long nativeTextureLoadPixels(const void* pixels, unsigned long width,
                             unsigned long height, long stride) {
    // Same store, same handles, same lifetime rules as an encoded texture. The
    // only difference is where the decoded image came from.
    const bb::oart::CreatedImage created = bb::oart::CreateCachedImageFromPixels(
        pixels, static_cast<std::uint32_t>(width), static_cast<std::uint32_t>(height),
        static_cast<std::int32_t>(stride));
    try {
        return TextureStore::Instance().Add(
            created.cached, created.image,
            static_cast<std::size_t>(stride) * height);
    } catch (...) {
        bb::oart::ReleaseIntrusive(created.image);
        bb::oart::ReleaseIntrusive(created.cached);
        throw;
    }
}

void nativeTextureApply(IDispatch* fill, long handle) {
    NativeTexture& texture = TextureStore::Instance().Get(handle);
    // Re-resolved for every apply and never cached: a Shape deleted through
    // public COM still passes every pointer and vtable check in this chain.
    const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill);
    const bb::oart::ApplyFunctions functions =
        bb::oart::ResolveApplyFunctions(target.oartBase);
    bb::oart::ApplyCachedImage(functions, target, texture.cached());
    texture.countApply();
}

void nativeTextureRelease(long handle) {
    TextureStore::Instance().Release(handle);
}

void nativeTextureClear() noexcept {
    TextureStore::Instance().Clear();
}

bool nativeTextureOwnsHandle(long handle) {
    return handle >= kNativeHandleBase;
}

long nativeTextureCount() {
    return static_cast<long>(TextureStore::Instance().Count());
}

std::wstring nativeTextureReport(long handle) {
    const TextureStore& store = TextureStore::Instance();
    std::wostringstream out;
    out << L"textures=" << store.Count() << L";nextHandle=" << store.NextHandle()
        << L";creations=" << store.Creations() << L';';
    if (handle > 0) {
        const NativeTexture& texture = TextureStore::Instance().Get(handle);
        out << L"handle=" << handle << L";cached=0x" << std::hex
            << reinterpret_cast<std::uintptr_t>(texture.cached()) << std::dec
            << L";cachedCount=" << bb::oart::IntrusiveCount(texture.cached())
            << L";bytes=" << texture.byteCount() << L";applies=" << texture.applyCount()
            << L';';
    }
    return out.str();
}
