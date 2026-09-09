/**
 * @file picture_cache.cpp
 * Implementation of the one-call picture fill and its caches.
 *
 * ## The dispatch decision
 *
 * Native where the Shape class has a validated path, ordinary
 * `Fill.UserPicture` where it does not, a specific refusal where neither
 * applies. The class list is `requireFillableShapeClass`, which is also what the
 * raw texture API uses, so the two can never disagree.
 *
 * The distinction that matters: **falling back is a statement about the Shape
 * class, not a response to failure.** If the native path is refused because the
 * class has no receiver chain, the fallback is right. If it fails because the
 * Office build is not validated, because the Shape was deleted, or because the
 * image will not decode, that is returned as itself. Retrying such a failure on
 * a slower path would hide exactly the problems worth knowing about, and would
 * turn "your Office build is unsupported" into "everything is a bit slow".
 *
 * ## Why the Shape cache is keyed the way it is
 *
 * A Shape pointer is useless as a key: it can be freed and reused, and two
 * different Shapes at different times can share an address. `Shape.Id` alone is
 * no better - it is unique within a slide, not across a presentation, and
 * certainly not across open documents.
 *
 * The key is therefore a composite read out of the object itself: the
 * presentation's identity, the slide's `SlideID`, and the Shape's `Id`. Reading
 * it costs three Automation property fetches at about a microsecond each,
 * against roughly 190 microseconds for the apply it may avoid.
 *
 * A Shape that cannot produce a full key - a group child, whose parent chain
 * differs - is simply not cached. It still gets a correct apply; it just pays
 * for it. Refusing to guess at a key is the whole point.
 *
 * ## What the skip can and cannot see
 *
 * Before skipping, `Fill.Type` is checked to still be a picture fill, so a Shape
 * whose fill was cleared or replaced with a solid colour is re-applied. What the
 * check cannot see is a fill replaced with a *different picture*: that still
 * reads as a picture fill. Detecting it would mean comparing image identity on
 * every call, which costs more than the apply it saves.
 *
 * That is a documented limit with a documented remedy - `BB_InvalidateShape` -
 * rather than a silent hazard. See docs/picture_cache.md.
 */

#include "picture_cache.hpp"

#include "native_texture.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <cstring>
#include <filesystem>
#include <fstream>
#include <map>
#include <sstream>
#include <vector>

namespace bb::office {
namespace {

/// `Fill.Type` for a picture fill; the value a successful apply produces.
constexpr long kPictureFill = 6;

/**
 * Identifies a file well enough that editing it produces a different texture.
 *
 * The path alone would serve a stale image after the file changed on disk. Size
 * and last-write time make that impossible for any edit that changes either,
 * which is every edit a normal tool performs, and the check is one
 * GetFileAttributesExW - about two microseconds.
 */
struct FileKey {
    std::wstring path;
    std::uint64_t size = 0;
    std::uint64_t written = 0;

    bool operator<(const FileKey& other) const {
        if (path != other.path) {
            return path < other.path;
        }
        if (size != other.size) {
            return size < other.size;
        }
        return written < other.written;
    }
};

FileKey DescribeFile(const std::wstring& path) {
    WIN32_FILE_ATTRIBUTE_DATA attributes{};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &attributes)) {
        throw bb::Error(bb::BB_E_IMAGE_FILE_MISSING,
                        "Cannot read the image file at the path given");
    }
    if (attributes.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        throw bb::Error(bb::BB_E_IMAGE_FILE_MISSING, "The path given is a directory, not an image");
    }
    FileKey key;
    key.path = path;
    key.size = (static_cast<std::uint64_t>(attributes.nFileSizeHigh) << 32) |
               attributes.nFileSizeLow;
    key.written = (static_cast<std::uint64_t>(attributes.ftLastWriteTime.dwHighDateTime) << 32) |
                  attributes.ftLastWriteTime.dwLowDateTime;
    return key;
}

/**
 * A Shape's identity as values rather than as a pointer.
 *
 * `valid` is false when any part could not be read; such a Shape is applied to
 * but never cached, because a partial key would collide with other Shapes.
 */
struct ShapeKey {
    long presentation = 0;
    long slide = 0;
    long shape = 0;
    bool valid = false;

    bool operator<(const ShapeKey& other) const {
        if (presentation != other.presentation) {
            return presentation < other.presentation;
        }
        if (slide != other.slide) {
            return slide < other.slide;
        }
        return shape < other.shape;
    }
};

/**
 * Reads a Shape's composite identity.
 *
 * Never throws: an unkeyable Shape is a reason to skip the cache, not to fail an
 * apply that would otherwise have worked.
 */
ShapeKey DescribeShape(IDispatch* shape) noexcept {
    ShapeKey key;
    try {
        key.shape = bb::get(shape, L"Id").integer();
        bb::Value parent = bb::get(shape, L"Parent");
        if (parent.v.vt != VT_DISPATCH || !parent.obj()) {
            return key;
        }
        // Shape.Parent is the Slide for a top-level Shape. For a group child it
        // is the group, which has no SlideID - so those go uncached rather than
        // being given a key that means something different.
        key.slide = bb::get(parent.obj(), L"SlideID").integer();
        bb::Value presentation = bb::get(parent.obj(), L"Parent");
        if (presentation.v.vt != VT_DISPATCH || !presentation.obj()) {
            return key;
        }
        // Presentations have no numeric id, but every open one has a distinct
        // window-independent hash of its full name plus its index. The index
        // alone would shift as documents open and close.
        bb::Value name = bb::get(presentation.obj(), L"FullName");
        std::wstring text;
        if (name.v.vt == VT_BSTR && name.v.bstrVal) {
            text.assign(name.v.bstrVal, SysStringLen(name.v.bstrVal));
        }
        if (text.empty()) {
            // An unsaved presentation has no FullName. Its Name ("Presentation1")
            // is unique among open documents, which is all this key needs.
            bb::Value shortName = bb::get(presentation.obj(), L"Name");
            if (shortName.v.vt == VT_BSTR && shortName.v.bstrVal) {
                text.assign(shortName.v.bstrVal, SysStringLen(shortName.v.bstrVal));
            }
        }
        if (text.empty()) {
            return key;
        }
        // FNV-1a over the document's name. Unsigned throughout: the constants
        // do not fit a 32-bit long, and signed overflow would be undefined.
        std::uint32_t hash = 2166136261u;
        for (wchar_t character : text) {
            hash ^= static_cast<std::uint32_t>(character);
            hash *= 16777619u;
        }
        key.presentation = static_cast<long>(hash);
        key.valid = true;
    } catch (...) {
        key.valid = false;
    }
    return key;
}

std::vector<std::uint8_t> ReadFile(const std::wstring& path) {
    // std::filesystem::path is what carries a wide path into ifstream portably;
    // the wstring overload is not standard.
    std::ifstream file(std::filesystem::path(path), std::ios::binary);
    if (!file) {
        throw bb::Error(bb::BB_E_IMAGE_FILE_MISSING, "Cannot open the image file at the path given");
    }
    std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(file)),
                                    std::istreambuf_iterator<char>());
    if (bytes.empty()) {
        throw bb::Error(bb::BB_E_INVALID_IMAGE, "The image file is empty");
    }
    return bytes;
}

/// Loads bytes into a native texture. The store copies them, so the array only
/// has to outlive the call.
long LoadTextureFromBytes(const std::vector<std::uint8_t>& bytes) {
    SAFEARRAYBOUND bound{static_cast<ULONG>(bytes.size()), 0};
    SAFEARRAY* array = SafeArrayCreate(VT_UI1, 1, &bound);
    if (!array) {
        throw std::bad_alloc();
    }
    struct Destroy {
        SAFEARRAY* value;
        ~Destroy() { SafeArrayDestroy(value); }
    } destroy{array};

    void* raw = nullptr;
    bb::check(SafeArrayAccessData(array, &raw), "SafeArrayAccessData");
    std::memcpy(raw, bytes.data(), bytes.size());
    SafeArrayUnaccessData(array);
    return nativeTextureLoad(array);
}

/**
 * Both caches and the counter, for the one apartment that owns them.
 *
 * A singleton is the honest model: the texture store beneath is one too, and
 * everything here is STA-bound.
 */
class PictureCache {
public:
    static PictureCache& Instance() {
        static PictureCache cache;
        return cache;
    }

    /// The texture for a file, decoding it only the first time it is seen.
    long TextureFor(const std::wstring& path) {
        const FileKey key = DescribeFile(path);
        const auto found = textures_.find(key);
        if (found != textures_.end()) {
            return found->second;
        }
        const long handle = LoadTextureFromBytes(ReadFile(path));
        // A file whose size or timestamp changed lands on a different key, so the
        // superseded entry for the same path is released rather than left to
        // accumulate across a session that keeps rewriting one file.
        DropOtherVersionsOf(path);
        textures_.emplace(key, handle);
        return handle;
    }

    /// True when @p shape already carries @p handle and still looks like it.
    bool AlreadyApplied(const ShapeKey& key, long handle, IDispatch* shape) {
        if (!key.valid) {
            return false;
        }
        const auto found = shapes_.find(key);
        if (found == shapes_.end() || found->second != handle) {
            return false;
        }
        // The record says this texture is already on this Shape. Confirm the
        // fill is still a picture before trusting it: a fill cleared or replaced
        // with a colour elsewhere must be re-applied, not skipped.
        try {
            bb::Value fill = bb::get(shape, L"Fill");
            if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
                return false;
            }
            if (bb::get(fill.obj(), L"Type").integer() != kPictureFill) {
                shapes_.erase(found);
                return false;
            }
        } catch (...) {
            shapes_.erase(found);
            return false;
        }
        ++skipped_;
        return true;
    }

    void Remember(const ShapeKey& key, long handle) {
        if (key.valid) {
            shapes_[key] = handle;
        }
    }

    void Forget(const ShapeKey& key) {
        if (key.valid) {
            shapes_.erase(key);
        }
    }

    void Clear() noexcept {
        for (const auto& [key, handle] : textures_) {
            try {
                nativeTextureRelease(handle);
            } catch (...) {
                // Teardown: a handle the store already dropped is not a failure.
            }
        }
        textures_.clear();
        shapes_.clear();
    }

    PictureCacheStats Stats() const noexcept {
        return PictureCacheStats{textures_.size(), shapes_.size(), skipped_};
    }

private:
    PictureCache() = default;

    /**
     * Releases every texture held for @p path under an older file version.
     *
     * Called just before the new version is inserted, so every existing entry
     * for that path is by definition stale - the file changed size or timestamp,
     * which is what put it on a different key.
     */
    void DropOtherVersionsOf(const std::wstring& path) {
        for (auto entry = textures_.begin(); entry != textures_.end();) {
            if (entry->first.path != path) {
                ++entry;
                continue;
            }
            // Every Shape remembered as carrying it must forget it too, or the
            // apply of the new version would be skipped as redundant.
            const long stale = entry->second;
            for (auto shape = shapes_.begin(); shape != shapes_.end();) {
                if (shape->second == stale) {
                    shape = shapes_.erase(shape);
                } else {
                    ++shape;
                }
            }
            try {
                nativeTextureRelease(stale);
            } catch (...) {
                // The store may already have dropped it; not a failure here.
            }
            entry = textures_.erase(entry);
        }
    }

    std::map<FileKey, long> textures_;
    std::map<ShapeKey, long> shapes_;
    std::uint64_t skipped_ = 0;
};

/// The ordinary Office route, for classes with no validated native path.
void ApplyThroughUserPicture(IDispatch* shape, const std::wstring& path) {
    try {
        bb::Value fill = bb::get(shape, L"Fill");
        if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
            throw bb::Error(bb::BB_E_SHAPE_CLASS_UNSUPPORTED, "Shape has no Fill to set");
        }
        bb::call(fill.obj(), L"UserPicture", {bb::Value(path.c_str())});
    } catch (const bb::Error& error) {
        std::ostringstream out;
        out << "This Shape class has no native picture-fill path, and Office's own "
               "Fill.UserPicture refused it too: "
            << error.what();
        throw bb::Error(bb::BB_E_FALLBACK_REFUSED, out.str());
    }
}

} // namespace

void ApplyPictureCached(IDispatch* shape, const std::wstring& path) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    if (path.empty()) {
        throw bb::Error(E_INVALIDARG, "An image path is required");
    }

    /*
     * Which route to take is decided from the Shape's *class* alone, before any
     * work happens. A class with no native path falls back; a class with one
     * gets the native apply and keeps whatever error it produces.
     */
    bool nativeClass = true;
    std::string classRefusal;
    try {
        requireFillableShapeClass(shape);
    } catch (const bb::Error& error) {
        nativeClass = false;
        classRefusal = error.what();
    }

    if (!nativeClass) {
        ApplyThroughUserPicture(shape, path);
        return;
    }

    PictureCache& cache = PictureCache::Instance();
    const long handle = cache.TextureFor(path);
    const ShapeKey key = DescribeShape(shape);
    if (cache.AlreadyApplied(key, handle, shape)) {
        return;
    }

    bb::Value fill = bb::get(shape, L"Fill");
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        throw bb::Error(E_INVALIDARG, "Shape has no Fill object");
    }
    // Any failure from here is reported as itself. A deleted Shape, an
    // unvalidated build or a corrupt image are all things the caller needs to
    // see, not things to retry more slowly.
    nativeTextureApply(fill.obj(), handle);
    cache.Remember(key, handle);
}

void InvalidateShapeCache(IDispatch* shape) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    PictureCache::Instance().Forget(DescribeShape(shape));
}

void ClearPictureCache() noexcept {
    PictureCache::Instance().Clear();
}

PictureCacheStats GetPictureCacheStats() noexcept {
    return PictureCache::Instance().Stats();
}

} // namespace bb::office
