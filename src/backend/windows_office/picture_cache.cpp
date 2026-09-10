/**
 * @file picture_cache.cpp
 * Implementation of the one-call picture fill and its caches.
 *
 * ## The dispatch decision
 *
 * Native where the Shape class has a validated path, ordinary
 * `Fill.UserPicture` where it does not, a specific refusal where neither
 * applies. The decision comes from `ClassifyShapeForNativePictureFill`, the one
 * semantic authority, which the raw texture API also asks - so the two surfaces
 * can never disagree about a Shape.
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
 * A Shape that cannot produce a full key is simply not cached. It still gets a
 * correct apply; it just pays for it. Refusing to guess at a key is the point.
 *
 * Group children *can* be keyed - `tools/probe_shape_identity.ps1` measured a
 * child's `Parent` to be the *Slide*, not the group, so it reports a `SlideID`,
 * and its `Id` collides with nothing on the slide. They are still deliberately
 * **not** cached, for a different reason, given below.
 *
 * ## Groups are never cached, and that is not conservatism
 *
 * Filling a group changes what its children render.
 * `tools/probe_group_fill_propagation.ps1` renders a child to PNG before and
 * after the group is filled and compares the bytes: 33,515 against 35,725, not
 * equal. `Fill.Type` stays 6 through all of it, so no cheap property reveals it.
 *
 * That makes a child's remembered texture stale the moment its group is filled.
 * A later request to put the child's own image back would be skipped as
 * redundant and the child would keep showing the group's image - a wrong picture,
 * silently, which is the one failure this cache must never produce. The reverse
 * holds too: filling a child changes what the group displays, so a group's
 * remembered texture goes stale when any child is filled.
 *
 * Tracking that properly would mean invalidating a whole subtree on every group
 * apply and every child apply, in both directions, including nested groups. The
 * cheap, obviously-correct rule is to not cache either: a group and anything
 * inside one always does real work. Groups are a small share of applies, and one
 * extra property read is the entire cost of being sure.
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
#include "shape_policy.hpp"

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
    key.size =
        (static_cast<std::uint64_t>(attributes.nFileSizeHigh) << 32) | attributes.nFileSizeLow;
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
    /// False when this Shape must not be cached at all - see DescribeShape.
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
ShapeKey DescribeShape(IDispatch* shape, long shapeType) noexcept {
    ShapeKey key;
    // A group's fill propagates to its children, and a child's fill changes what
    // the group shows, so a remembered texture on either side goes stale when the
    // other is filled. Neither is cached; see the file comment for the pixel
    // evidence. msoGroup is 6.
    constexpr long kGroup = 6;
    if (shapeType == kGroup) {
        return key;
    }
    try {
        // ParentGroup answers only for a Shape inside a group. One property read,
        // about a microsecond, against the 190 an incorrect skip would misplace.
        bb::Value owner = bb::get(shape, L"ParentGroup");
        if (owner.v.vt == VT_DISPATCH && owner.obj()) {
            return key;
        }
    } catch (const bb::Error&) {
        // Top-level Shapes refuse the question, which is the common case and
        // means exactly what it should: this Shape is not inside a group.
    }
    try {
        key.shape = bb::get(shape, L"Id").integer();
        bb::Value parent = bb::get(shape, L"Parent");
        if (parent.v.vt != VT_DISPATCH || !parent.obj()) {
            return key;
        }
        // Shape.Parent is the Slide. Anything that does not report a SlideID
        // goes uncached rather than being given a key that means something else.
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
        throw bb::Error(bb::BB_E_IMAGE_FILE_MISSING,
                        "Cannot open the image file at the path given");
    }
    std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(file)),
                                    std::istreambuf_iterator<char>());
    if (bytes.empty()) {
        throw bb::Error(bb::BB_E_INVALID_IMAGE, "The image file is empty");
    }
    return bytes;
}

/**
 * Decodes bytes into an image this cache owns outright.
 *
 * Deliberately *not* a public texture handle. The cache used to take one, which
 * meant `ClearTextures()` - or a caller releasing that particular handle -
 * destroyed the image the cache was still pointing at, and the next
 * `UserPicture2` failed with "handle ... is not valid (it was released)". The
 * cache now holds its own reference, so the two lifetimes are independent.
 *
 * The store copies the bytes, so the array only has to outlive the call.
 */
TextureRef LoadTextureFromBytes(const std::vector<std::uint8_t>& bytes) {
    SAFEARRAYBOUND bound{static_cast<ULONG>(bytes.size()), 0};
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
    bb::check(SafeArrayAccessData(array, &raw), "SafeArrayAccessData");
    std::memcpy(raw, bytes.data(), bytes.size());
    SafeArrayUnaccessData(array);
    return CreateTextureFromBytes(array);
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

    /// The image for a file, decoding it only the first time it is seen.
    TextureRef TextureFor(const std::wstring& path) {
        const FileKey key = DescribeFile(path);
        const auto found = textures_.find(key);
        if (found != textures_.end()) {
            return found->second;
        }
        TextureRef texture = LoadTextureFromBytes(ReadFile(path));
        // A file whose size or timestamp changed lands on a different key, so the
        // superseded entry for the same path is dropped rather than left to
        // accumulate across a session that keeps rewriting one file.
        DropOtherVersionsOf(path);
        textures_.emplace(key, texture);
        return texture;
    }

    /// True when @p shape already carries this image and still looks like it.
    bool AlreadyApplied(const ShapeKey& key, std::uint64_t textureId, IDispatch* shape) {
        if (!key.valid) {
            return false;
        }
        const auto found = shapes_.find(key);
        if (found == shapes_.end() || found->second != textureId) {
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

    void Remember(const ShapeKey& key, std::uint64_t textureId) {
        if (key.valid) {
            shapes_[key] = textureId;
        }
    }

    void Forget(const ShapeKey& key) {
        if (key.valid) {
            shapes_.erase(key);
        }
    }

    /**
     * Drops every image this cache owns, and every Shape it remembers.
     *
     * It touches nothing the caller owns: an image that also has a public
     * texture handle stays alive behind that handle. Symmetrically,
     * `ClearTextures()` cannot reach these.
     */
    void Clear() noexcept {
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
            const std::uint64_t stale = TextureIdOf(entry->second);
            for (auto shape = shapes_.begin(); shape != shapes_.end();) {
                if (shape->second == stale) {
                    shape = shapes_.erase(shape);
                } else {
                    ++shape;
                }
            }
            // Erasing drops this cache's reference. If a public handle also
            // holds the image, it stays alive there - which is correct: the
            // caller asked for that handle and has not released it.
            entry = textures_.erase(entry);
        }
    }

    std::map<FileKey, TextureRef> textures_;
    std::map<ShapeKey, std::uint64_t> shapes_;
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
     * The route is chosen from the Shape's *class*, before any work happens, and
     * from an explicit verdict rather than from whether something threw.
     *
     * That distinction is the whole point. Treating any refusal as "fall back"
     * would route a Connector - which Office also refuses - and, worse, a deleted
     * or unusable Shape into the slower path, turning a real failure into a
     * confusing one. Only FallbackSupported falls back.
     */
    const office::ShapeClassification classification =
        office::ClassifyShapeForNativePictureFill(shape);
    switch (classification.eligibility) {
    case office::ShapeEligibility::FallbackSupported:
        ApplyThroughUserPicture(shape, path);
        return;
    case office::ShapeEligibility::Invalid:
        throw bb::Error(E_INVALIDARG, classification.reason);
    case office::ShapeEligibility::Unsupported:
        throw bb::Error(bb::BB_E_SHAPE_CLASS_UNSUPPORTED, classification.reason);
    case office::ShapeEligibility::NativeSupported:
        break;
    }

    PictureCache& cache = PictureCache::Instance();
    const TextureRef texture = cache.TextureFor(path);
    // The type was already read by the classifier; passing it on saves a second
    // Automation fetch on every apply.
    const ShapeKey key = DescribeShape(shape, classification.shapeType);
    if (cache.AlreadyApplied(key, TextureIdOf(texture), shape)) {
        return;
    }

    bb::Value fill = bb::get(shape, L"Fill");
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        throw bb::Error(E_INVALIDARG, "Shape has no Fill object");
    }
    // Any failure from here is reported as itself. A deleted Shape, an
    // unvalidated build or a corrupt image are all things the caller needs to
    // see, not things to retry more slowly.
    ApplyTextureRef(fill.obj(), texture);
    cache.Remember(key, TextureIdOf(texture));
}

void InvalidateShapeCache(IDispatch* shape) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    const ShapeClassification classification = ClassifyShapeForNativePictureFill(shape);
    PictureCache::Instance().Forget(DescribeShape(shape, classification.shapeType));
}

void ClearPictureCache() noexcept {
    PictureCache::Instance().Clear();
}

PictureCacheStats GetPictureCacheStats() noexcept {
    return PictureCache::Instance().Stats();
}

} // namespace bb::office
