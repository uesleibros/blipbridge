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

#include "apply_skip_cache.hpp"

#include "native_apply.hpp"
#include "oart_layout.hpp"
#include "shape_policy.hpp"

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

} // namespace

namespace bb::office {

/**
 * One decoded image, owning one reference to each of the GFX objects the
 * creator returned.
 *
 * Lives outside the anonymous namespace because ownership is now shared: the
 * public handle table and the picture cache each hold a TextureRef to the same
 * instance, and the image dies only when the last of them lets go.
 */
class NativeTexture {
  public:
    NativeTexture(void* cached, void* image, std::size_t byteCount)
        : cached_(cached), image_(image), byteCount_(byteCount), id_(++nextId_) {}

    NativeTexture(const NativeTexture&) = delete;
    NativeTexture& operator=(const NativeTexture&) = delete;

    ~NativeTexture() {
        // Order mirrors creation: the image is the companion output, the cached
        // image is the one Office holds while a fill is in place.
        bb::oart::ReleaseIntrusive(image_);
        bb::oart::ReleaseIntrusive(cached_);
    }

    void* cached() const {
        return cached_;
    }

    std::size_t byteCount() const {
        return byteCount_;
    }

    unsigned long applyCount() const {
        return applyCount_;
    }

    void countApply() {
        ++applyCount_;
    }

    /// Process-unique, so the picture cache can say "the same image" without
    /// keeping it alive or dereferencing something already freed.
    std::uint64_t id() const {
        return id_;
    }

  private:
    void* cached_ = nullptr;
    void* image_ = nullptr;
    std::size_t byteCount_ = 0;
    unsigned long applyCount_ = 0;
    std::uint64_t id_ = 0;

    static std::uint64_t nextId_;
};

std::uint64_t NativeTexture::nextId_ = 0;

} // namespace bb::office

namespace {

using bb::office::NativeTexture;
using bb::office::TextureRef;

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

    /**
     * Gives @p texture a public handle, adding this table as a second owner.
     *
     * The table used to own the image outright, which is what let ClearTextures
     * destroy an image the picture cache was still using.
     */
    long Add(TextureRef texture) {
        RequireOwningThread(true);
        const long handle = nextHandle_++;
        textures_.emplace(handle, std::move(texture));
        return handle;
    }

    void CountCreation() {
        ++creations_;
    }

    /**
     * Why a handle is not usable, told apart rather than lumped together.
     *
     * A handle that was issued and released is a different mistake from one that
     * was never issued - a lifetime bug against a typo - and the counter only
     * ever rises, so the store can say which with certainty. The portable
     * backend answers the same question the same way; a caller should not be
     * able to tell which backend refused them.
     */
    std::string DescribeBadHandle(long handle) const {
        std::ostringstream out;
        out << "Texture handle " << handle << " is not valid (";
        out << (handle >= kNativeHandleBase && handle < nextHandle_
                    ? "it was released; handles are never recycled"
                    : "it was never created");
        out << ")";
        return out.str();
    }

    const TextureRef& GetRef(long handle) {
        RequireOwningThread(false);
        const auto found = textures_.find(handle);
        if (found == textures_.end()) {
            throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, DescribeBadHandle(handle));
        }
        return found->second;
    }

    NativeTexture& Get(long handle) {
        RequireOwningThread(false);
        return *GetRef(handle);
    }

    void Release(long handle) {
        RequireOwningThread(false);
        if (textures_.erase(handle) == 0) {
            throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, DescribeBadHandle(handle));
        }
    }

    /**
     * Drops every public handle. Safe to call when the store was never used.
     *
     * This releases the *handles*, not necessarily the images: one the picture
     * cache also holds stays alive and usable, which is the whole point of the
     * split. Deliberately not thread-checked, because this also runs from
     * teardown paths.
     */
    void Clear() noexcept {
        textures_.clear();
    }

    std::size_t Count() const {
        return textures_.size();
    }

    long NextHandle() const {
        return nextHandle_;
    }

    /// Total decodes since the process started. A reuse claim is only credible
    /// if this stays far below the number of applies.
    unsigned long Creations() const {
        return creations_;
    }

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

    std::map<long, TextureRef> textures_;
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

namespace bb::office {

/// Wraps a freshly created image, releasing its GFX references if the wrap
/// itself throws - the only window in which nothing else owns them yet.
static TextureRef Own(const bb::oart::CreatedImage& created, std::size_t byteCount) {
    try {
        TextureRef texture =
            std::make_shared<NativeTexture>(created.cached, created.image, byteCount);
        TextureStore::Instance().CountCreation();
        return texture;
    } catch (...) {
        bb::oart::ReleaseIntrusive(created.image);
        bb::oart::ReleaseIntrusive(created.cached);
        throw;
    }
}

TextureRef CreateTextureFromBytes(SAFEARRAY* bytes) {
    LONG lower = 0;
    LONG upper = 0;
    if (!bytes || SafeArrayGetDim(bytes) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional Byte array");
    }
    bb::check(SafeArrayGetLBound(bytes, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(bytes, 1, &upper), "SafeArrayGetUBound");

    // Decoding is the expensive half, and it happens exactly once per image.
    return Own(bb::oart::CreateCachedImageFromBytes(bytes),
               static_cast<std::size_t>(upper - lower) + 1);
}

TextureRef CreateTextureFromPixels(const void* pixels,
                                   unsigned long width,
                                   unsigned long height,
                                   long stride) {
    // Same lifetime rules as an encoded image; only the decode differs.
    return Own(bb::oart::CreateCachedImageFromPixels(pixels,
                                                     static_cast<std::uint32_t>(width),
                                                     static_cast<std::uint32_t>(height),
                                                     static_cast<std::int32_t>(stride)),
               static_cast<std::size_t>(stride) * height);
}

long RegisterTexture(TextureRef texture) {
    if (!texture) {
        throw bb::Error(E_POINTER, "Cannot register a null texture");
    }
    return TextureStore::Instance().Add(std::move(texture));
}

void ApplyTextureRef(IDispatch* fill, const TextureRef& texture) {
    if (!texture) {
        throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, "No texture to apply");
    }
    // Re-resolved for every apply and never cached: a Shape deleted through
    // public COM still passes every pointer and vtable check in this chain.
    const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill);
    const bb::oart::ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);
    bb::oart::ApplyCachedImage(functions, target, texture->cached());
    texture->countApply();
}

TextureRef LookupTexture(long handle) {
    return TextureStore::Instance().GetRef(handle);
}

void* CachedImageOf(const TextureRef& texture) {
    return texture ? texture->cached() : nullptr;
}

std::uint64_t TextureIdOf(const TextureRef& texture) {
    return texture ? texture->id() : 0;
}

void CountTextureApplies(const TextureRef& texture, unsigned long times) {
    if (!texture) {
        return;
    }
    for (unsigned long index = 0; index < times; ++index) {
        texture->countApply();
    }
}

} // namespace bb::office

long nativeTextureLoad(SAFEARRAY* bytes) {
    return bb::office::RegisterTexture(bb::office::CreateTextureFromBytes(bytes));
}

long nativeTextureLoadPixels(const void* pixels,
                             unsigned long width,
                             unsigned long height,
                             long stride) {
    return bb::office::RegisterTexture(
        bb::office::CreateTextureFromPixels(pixels, width, height, stride));
}

void nativeTextureApply(IDispatch* fill, long handle) {
    // The handle is resolved to a reference first, so the apply itself is the
    // same code the picture cache runs.
    bb::office::ApplyTextureRef(fill, TextureStore::Instance().GetRef(handle));
}

void nativeTextureRelease(long handle) {
    // Forget the image *before* the handle goes, while its id is still readable.
    // If another owner keeps the image alive this costs one apply; if this was
    // the last owner it prevents a stale claim outliving it.
    std::uint64_t id = 0;
    try {
        id = bb::office::TextureIdOf(TextureStore::Instance().GetRef(handle));
    } catch (...) {
        // An unknown handle is reported by Release below, not here.
    }
    TextureStore::Instance().Release(handle);
    if (id != 0) {
        bb::office::ForgetTexture(id);
    }
}

void nativeTextureApplyToShape(IDispatch* shape, long handle, long shapeType) {
    const bb::office::TextureRef texture = TextureStore::Instance().GetRef(handle);
    bb::office::ApplyTextureRef(bb::get(shape, L"Fill").obj(), texture);

    /*
     * An apply that changes a fill without saying so leaves a stale claim behind
     * it, and a later skip would honour that claim and show the wrong picture.
     * So this path writes the record too, even though it never reads it.
     *
     * The test comes first because the Shape key costs four Automation fetches.
     * Nothing is remembered until something asks for a skip, so a caller who
     * only ever calls ApplyTexture pays one comparison and no fetches.
     */
    if (bb::office::AnyRemembered()) {
        bb::office::RememberApplied(bb::office::DescribeShape(shape, shapeType),
                                    bb::office::TextureIdOf(texture));
    }
}

void nativeTextureApplyIfChanged(IDispatch* shape, long handle, long shapeType, bool* skipped) {
    if (skipped) {
        *skipped = false;
    }
    // Resolved first, and by reference: an unknown handle must fail here exactly
    // as nativeTextureApply would, before any decision about skipping.
    const bb::office::TextureRef texture = TextureStore::Instance().GetRef(handle);
    const std::uint64_t id = bb::office::TextureIdOf(texture);
    const bb::office::ShapeKey key = bb::office::DescribeShape(shape, shapeType);

    if (bb::office::AlreadyCarries(key, id, shape)) {
        if (skipped) {
            *skipped = true;
        }
        return;
    }

    // Only now is Fill worth fetching: the skip check reads the Shape's fill
    // itself, so fetching up front would charge the cheap path for the dear one.
    bb::office::ApplyTextureRef(bb::get(shape, L"Fill").obj(), texture);
    bb::office::RememberApplied(key, id);
}

void nativeTextureClear() noexcept {
    // Dropping every handle may drop the last reference to some images, and a
    // Shape must not keep claiming one that no longer exists - a later image
    // could otherwise inherit the claim and be wrongly skipped.
    bb::office::ForgetAllApplied();
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
    out << L"textures=" << store.Count() << L";nextHandle=" << store.NextHandle() << L";creations="
        << store.Creations() << L';';
    if (handle > 0) {
        const NativeTexture& texture = TextureStore::Instance().Get(handle);
        out << L"handle=" << handle << L";cached=0x" << std::hex
            << reinterpret_cast<std::uintptr_t>(texture.cached()) << std::dec << L";cachedCount="
            << bb::oart::IntrusiveCount(texture.cached()) << L";bytes=" << texture.byteCount()
            << L";applies=" << texture.applyCount() << L';';
    }
    return out.str();
}
