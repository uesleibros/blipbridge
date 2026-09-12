/**
 * @file portable_office_backend.cpp
 * The portable PowerPoint backend: everything the public API promises, driven
 * entirely through documented Office Automation.
 *
 * ## What this is for
 *
 * The accelerated backend reaches into OART and GFX. That work was derived
 * against one 64-bit Office build, and none of it transfers to 32-bit
 * PowerPoint: different calling convention, different module RVAs, different
 * object layouts, different record size. Recompiling it as x86 would produce
 * something that links and is wrong.
 *
 * So 32-bit PowerPoint had no backend at all, and BlipBridge refused every
 * texture call there. That is what this file changes. Nothing here is
 * reverse-engineered: it is `Shape.Fill.UserPicture` and `ShapeRange.Fill`,
 * both documented, both stable across Office versions, both architecture-neutral
 * by construction.
 *
 * ## What a caller gets, and what they give up
 *
 * Everything works: textures, the range apply, the skip cache, the whole image
 * pipeline - crop, orient, resize, quad warp - and `BB_ApplyPicture`. The
 * pipeline was always portable maths; only its last step needed a backend, and
 * now it has one.
 *
 * What is given up is speed, and the honest way to say it is that this is the
 * ordinary Office route with BlipBridge's image work in front of it:
 *
 *  - a texture is encoded to a temporary PNG once, and Office reads that file;
 *  - Office embeds a copy per fill rather than sharing one cached image;
 *  - there is no private transaction to skip, so an apply costs what Office
 *    charges for an apply.
 *
 * `BB_CAP_NATIVE_BACKEND` is therefore **not** reported here, and that is not a
 * formality. A caller who benchmarks against the accelerated path and finds it
 * slower should be able to discover why from the capability bits rather than by
 * guessing, so the bit that means "the accelerated backend is usable" says no.
 * Every bit that means "this feature works" says yes, because it does.
 *
 * ## What it shares with the accelerated backend, deliberately
 *
 * The Shape class policy, the Shape identity key and the skip cache are the same
 * code, not a second implementation. Those three decide which Shapes may be
 * filled, which may be cached, and when an apply may be skipped - and a
 * divergence there would be a behaviour difference between architectures with no
 * visible cause. They are portable Automation already; only the apply underneath
 * them differs.
 *
 * Threading: single-threaded-apartment, like everything else that touches these
 * objects.
 */

#include "backend.hpp"
#include "backend_guard.hpp"
#include "portable_office/portable_texture.hpp"
#include "windows_office/apply_skip_cache.hpp"
#include "windows_office/shape_identity.hpp"
#include "windows_office/shape_policy.hpp"

#include "../image/resample.hpp"
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <iterator>
#include <map>
#include <sstream>
#include <string>
#include <vector>

// MinGW requires Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

namespace bb {
namespace {

/**
 * Sets one Shape's fill from a file, through Office's own API.
 *
 * The single place this backend changes a document. Everything else in the file
 * decides *whether* to call it and *what file* to hand it.
 */
void FillFromPath(IDispatch* shape, const std::wstring& path) {
    Value fill = get(shape, L"Fill");
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        throw Error(BB_E_SHAPE_CLASS_UNSUPPORTED, "Shape has no Fill to set");
    }
    try {
        call(fill.obj(), L"UserPicture", {Value(path.c_str())});
        // Counted after it succeeded, so the number means "Office edited the
        // document", which is what a skip is measured against.
        portable::CountOfficeFill();
    } catch (const Error& error) {
        std::ostringstream out;
        out << "Office's Fill.UserPicture refused this Shape: " << error.what();
        throw Error(BB_E_FALLBACK_REFUSED, out.str());
    }
}

/**
 * What makes one image file distinct from another: its path, its size and when
 * it was last written.
 *
 * The path alone is not an identity. A caller who writes a new image over the
 * same temporary file - which is an entirely ordinary thing to do - would
 * otherwise be told the Shape already carries that picture, and the apply would
 * be skipped. The Shape would keep the *old* image, silently, and nothing would
 * indicate it. That is the one failure this whole mechanism exists to prevent,
 * so the key is the same composite the accelerated backend's cache uses.
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

/// Reads @p path's identity, refusing a path that is not a readable file.
FileKey DescribeFile(const std::wstring& path) {
    WIN32_FILE_ATTRIBUTE_DATA attributes{};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &attributes)) {
        std::ostringstream out;
        out << "Cannot read the image file at the path given (Windows error " << GetLastError()
            << ")";
        throw Error(BB_E_IMAGE_FILE_MISSING, out.str());
    }
    if ((attributes.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
        throw Error(BB_E_IMAGE_FILE_MISSING, "The path given is a directory, not an image");
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
 * The image identities behind `BB_ApplyPicture`.
 *
 * The accelerated backend caches a *decoded* image per file, because decoding is
 * the expensive part there and the apply hands Office a resource it already
 * holds. Here the apply hands Office a path, so the caller's own file is already
 * the best thing to give it: decoding it only to encode it again into a second
 * temporary file would cost two conversions and change the image.
 *
 * So what is kept here is only the *identity* of each file - the number the skip
 * cache compares - and the file itself is passed straight through. That keeps
 * `BB_ApplyPicture` and `ApplyTextureIfChanged` sharing one record of what each
 * Shape carries, which is what stops a stale skip showing the wrong picture.
 */
class PictureIdentities {
  public:
    static PictureIdentities& Instance() {
        static PictureIdentities identities;
        return identities;
    }

    std::uint64_t IdFor(const FileKey& key) {
        const auto entry = ids_.find(key);
        if (entry != ids_.end()) {
            return entry->second;
        }

        /*
         * A new identity for a path this already knows means the file was
         * rewritten, so the old identity describes an image that no longer
         * exists anywhere. Dropping it keeps this bounded for a caller who
         * writes over one scratch file in a loop - otherwise every rewrite
         * would leave an entry behind for the rest of the process.
         *
         * The map is ordered by path first, so one path's entries are
         * contiguous and the old ones are exactly the range in between.
         */
        const auto samePath = [&key](const auto& existing) {
            return existing.first.path == key.path;
        };
        for (auto it = ids_.begin(); it != ids_.end();) {
            it = samePath(*it) ? ids_.erase(it) : std::next(it);
        }

        // Ids come from the texture store's counter so a picture identity can
        // never collide with a texture identity in the shared skip cache.
        const std::uint64_t id = portable::NextSharedImageId();
        ids_.emplace(key, id);
        return id;
    }

    void Clear() noexcept {
        ids_.clear();
    }

    std::size_t Count() const noexcept {
        return ids_.size();
    }

  private:
    std::map<FileKey, std::uint64_t> ids_;
};

/**
 * Applies a texture to one Shape and keeps the shared per-Shape record truthful.
 *
 * @p skipped, when given, turns this into the conditional apply: the Office edit
 * is avoided when the Shape demonstrably already carries this image. When it is
 * null the apply always happens, and the record is still updated - otherwise a
 * later conditional apply could skip on the strength of an image this call
 * replaced.
 */
void ApplyToShape(IDispatch* shape,
                  const portable::TextureRef& texture,
                  long shapeType,
                  bool* skipped) {
    const std::uint64_t id = portable::TextureIdOf(texture);
    const office::ShapeKey key = office::DescribeShape(shape, shapeType);

    if (skipped) {
        *skipped = false;
        if (key.valid && office::AlreadyCarries(key, id, shape)) {
            *skipped = true;
            return;
        }
    }

    // Encoding happens here, on the first apply of this texture, not at load:
    // a texture that is created and never applied never touches the disk.
    FillFromPath(shape, texture->Path());
    texture->CountApplies(1);
    if (key.valid) {
        office::RememberApplied(key, id);
    }
}

/// How many members @p range holds, with a legible refusal for a non-range.
long MemberCount(IDispatch* range) {
    try {
        return get(range, L"Count").integer();
    } catch (const Error&) {
        // A Shape has no Count. Saying so beats "member 1 is refused".
        throw Error(E_INVALIDARG,
                    "This object is not a ShapeRange: it did not answer Count. Pass "
                    "Slide.Shapes.Range(...) or ShapeRange, not a single Shape - "
                    "BB_ApplyTexture takes a Shape.");
    }
}

/**
 * Classifies every member of @p range before anything is filled.
 *
 * All or nothing: one member the policy refuses refuses the whole range with
 * nothing applied. A range that filled some members and then failed would leave
 * the caller with a document they have to inspect to understand, and no way to
 * retry cleanly.
 *
 * ## Why a fallback-only class is refused here, when this backend could fill it
 *
 * Every route out of this backend ends in `Fill.UserPicture`, so a Table - which
 * has no native path but which Office's own fill accepts - would work perfectly
 * well in a range here. It is refused anyway, to match what the accelerated
 * backend does with the same range.
 *
 * That is a deliberate trade, and the reasoning is the one this file is built
 * on: `BB_ApplyTextureRange` has to mean one thing. A range that filled on
 * 32-bit and was refused on 64-bit would be a difference a caller discovers in
 * production, from a document that came out wrong, with nothing in the API to
 * explain it. Being uniformly stricter costs a capability nobody was promised;
 * being non-uniformly permissive costs trust in the whole surface.
 *
 * The same rule already governs the single-Shape apply, which refuses a Table on
 * both backends too. `BB_ApplyPicture` is the entry point that accepts fallback
 * classes, on both.
 */
std::vector<Value> RequireEveryMemberEligible(IDispatch* range, std::vector<long>& types) {
    const long count = MemberCount(range);
    if (count <= 0) {
        throw Error(E_INVALIDARG, "The ShapeRange holds no Shapes");
    }

    std::vector<Value> members;
    members.reserve(static_cast<std::size_t>(count));
    types.reserve(static_cast<std::size_t>(count));
    for (long index = 1; index <= count; ++index) {
        Value member = item(range, index);
        if (member.v.vt != VT_DISPATCH || !member.obj()) {
            std::ostringstream out;
            out << "Range member " << index << " of " << count << " is not a Shape";
            throw Error(E_INVALIDARG, out.str());
        }
        const office::ShapeClassification classification =
            office::ClassifyShapeForNativePictureFill(member.obj());
        if (!classification.native()) {
            std::ostringstream out;
            out << "Range member " << index << " of " << count
                << " is refused: " << classification.reason
                << ". The fill would reach every member, so the whole range is refused and "
                   "nothing has been applied.";
            throw Error(classification.eligibility == office::ShapeEligibility::Invalid
                            ? E_INVALIDARG
                            : BB_E_SHAPE_CLASS_UNSUPPORTED,
                        out.str());
        }
        types.push_back(classification.shapeType);
        members.push_back(std::move(member));
    }
    return members;
}

/**
 * The portable implementation.
 *
 * Holds no state of its own: the texture store and the skip cache are module
 * singletons, so the C ABI and anything else in the process address the same
 * textures and cannot disagree about what is loaded.
 */
class PortableOfficeBackend final : public Backend {
  public:
    const char* Name() const noexcept override {
        return "windows-office-portable";
    }

    BackendResult Probe() noexcept override {
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return BackendResult::Failure(
                BackendStatus::UnsupportedHost,
                "BlipBridge fills PowerPoint Shapes; this process is not PowerPoint");
        }
        return BackendResult::Success();
    }

    BackendCapabilities Capabilities() const noexcept override {
        BackendCapabilities capabilities;
        // Every capability here needs a live PowerPoint. Reporting one in a
        // process that is not PowerPoint would be a lie a caller could act on.
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return capabilities;
        }
        /*
         * nativeBackend stays false, and it is the one bit that is *about* this
         * backend rather than about a feature. Everything else is true because
         * the feature genuinely works; this one is false because the accelerated
         * path genuinely is not here, and a caller measuring performance should
         * be able to find that out by asking.
         */
        capabilities.nativeBackend = false;
        capabilities.applyPicture = true;
        capabilities.pickUpFallback = true;
        capabilities.memoryImage = true;
        capabilities.cachedTexture = true;
        capabilities.batchApply = true;
        capabilities.rangeApply = true;
        capabilities.imagePipeline = true;
        capabilities.rawPixels = true;
        capabilities.scaledPixels = true;
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
        return Guarded(
            [&] { *out = portable::RegisterTexture(portable::CreateTextureFromBytes(bytes, length)); });
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
        if (stride < static_cast<std::int32_t>(width) * 4) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Stride must be at least width*4 bytes for BGRA32");
        }
        return Guarded([&] {
            *out = portable::RegisterTexture(
                portable::CreateTextureFromPixels(pixels, width, height, stride));
        });
    }

    BackendResult ApplyTexture(void* shape, std::uint64_t texture) noexcept override {
        return Guarded([&] {
            long shapeType = 0;
            IDispatch* dispatch = RequireFillableShape(shape, &shapeType);
            ApplyToShape(dispatch, portable::LookupTexture(texture), shapeType, nullptr);
        });
    }

    BackendResult ApplyTextureIfChanged(void* shape,
                                        std::uint64_t texture,
                                        bool* skipped) noexcept override {
        if (skipped) {
            *skipped = false;
        }
        return Guarded([&] {
            // Same gate, same order as the ordinary apply: an ineligible Shape is
            // refused before the skip cache is consulted, so this cannot become a
            // way to reach a fill the ordinary apply would reject.
            long shapeType = 0;
            IDispatch* dispatch = RequireFillableShape(shape, &shapeType);
            ApplyToShape(dispatch, portable::LookupTexture(texture), shapeType, skipped);
        });
    }

    BackendResult ApplyTextureRange(void* shapeRange,
                                    std::uint64_t texture,
                                    std::uint32_t* applied) noexcept override {
        if (applied) {
            *applied = 0;
        }
        return Guarded([&] {
            // A ShapeRange is not a Shape, so the Shape-class gate cannot be asked
            // about it here: the pointer is checked, and every member is
            // classified inside, before anything is filled.
            IDispatch* dispatch = RequireDispatchShape(shapeRange);
            ReleaseDispatch release{dispatch};

            std::vector<long> types;
            const std::vector<Value> members = RequireEveryMemberEligible(dispatch, types);
            const portable::TextureRef image = portable::LookupTexture(texture);
            // Encode once, before the fill, so a failure to produce the file
            // refuses the range with nothing applied rather than part-way through.
            const std::wstring& path = image->Path();

            /*
             * One call on the range, not one per member. `ShapeRange.Fill` is a
             * real Office object that fans out to its members internally, which
             * is the same thing the accelerated backend's range receiver does -
             * so both architectures produce one Office edit, and one undo entry,
             * for a range apply.
             */
            FillFromPath(dispatch, path);
            image->CountApplies(static_cast<unsigned long>(members.size()));

            // The shared record has to learn what every member now carries, or a
            // later conditional apply would re-fill Shapes that are already right.
            const std::uint64_t id = portable::TextureIdOf(image);
            for (std::size_t index = 0; index < members.size(); ++index) {
                const office::ShapeKey key =
                    office::DescribeShape(members[index].obj(), types[index]);
                if (key.valid) {
                    office::RememberApplied(key, id);
                }
            }
            if (applied) {
                *applied = static_cast<std::uint32_t>(members.size());
            }
        });
    }

    BackendResult ReleaseTexture(std::uint64_t texture) noexcept override {
        return Guarded([&] {
            // Forget every Shape remembered as carrying this image, so a later
            // image cannot inherit a stale claim and be wrongly skipped.
            const std::uint64_t id = portable::TextureIdOf(portable::LookupTexture(texture));
            portable::ReleaseTexture(texture);
            office::ForgetTexture(id);
        });
    }

    void ClearTextures() noexcept override {
        portable::ClearTextures();
        office::ForgetAllApplied();
        // Every texture deleted its own file as it went, so this only takes the
        // directory they were in - and only if it is empty.
        portable::CleanUpTemporaryDirectory();
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
        // which one was wrong; repeating those checks here would only let the two
        // disagree.
        std::vector<std::uint8_t> scaled;
        const image::ResampleStatus status =
            image::Resample(pixels, width, height, stride, targetWidth, targetHeight,
                            static_cast<image::ScaleFilter>(filter), scaled);
        if (status != image::ResampleStatus::Ok) {
            const BackendStatus code = status == image::ResampleStatus::OutOfMemory
                                           ? BackendStatus::OutOfMemory
                                           : BackendStatus::InvalidArgument;
            return BackendResult::Failure(code, image::DescribeStatus(status));
        }
        return Guarded([&] {
            // The resampled buffer is tightly packed, so its stride is exactly
            // one row of BGRA.
            *out = portable::RegisterTexture(
                portable::CreateTextureFromPixels(scaled.data(),
                                                  targetWidth,
                                                  targetHeight,
                                                  static_cast<std::int32_t>(targetWidth) * 4));
        });
    }

    BackendResult ApplyPicture(void* shape, const std::uint16_t* path) noexcept override {
        if (!path || !*path) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "An image path is required");
        }
        return Guarded([&] {
            IDispatch* dispatch = RequireDispatchShape(shape);
            ReleaseDispatch release{dispatch};

            const std::wstring file(reinterpret_cast<const wchar_t*>(path));
            // Reading the identity is also how an unreadable path is refused,
            // so there is no separate existence check to fall out of step.
            const FileKey key = DescribeFile(file);

            /*
             * The route is chosen from the Shape's *class*, before any work
             * happens, and from an explicit verdict rather than from whether
             * something threw - the same rule the accelerated backend follows.
             * Both acceptable verdicts end in the same call here, but a
             * Connector must still be refused as unsupported rather than tried.
             */
            const office::ShapeClassification classification =
                office::ClassifyShapeForNativePictureFill(dispatch);
            switch (classification.eligibility) {
            case office::ShapeEligibility::Invalid:
                throw Error(E_INVALIDARG, classification.reason);
            case office::ShapeEligibility::Unsupported:
                throw Error(BB_E_SHAPE_CLASS_UNSUPPORTED, classification.reason);
            case office::ShapeEligibility::NativeSupported:
            case office::ShapeEligibility::FallbackSupported:
                break;
            }

            const std::uint64_t id = PictureIdentities::Instance().IdFor(key);
            const office::ShapeKey shapeKey =
                office::DescribeShape(dispatch, classification.shapeType);
            if (shapeKey.valid && office::AlreadyCarries(shapeKey, id, dispatch)) {
                return;
            }
            FillFromPath(dispatch, file);
            if (shapeKey.valid) {
                office::RememberApplied(shapeKey, id);
            }
        });
    }

    BackendResult InvalidateShape(void* shape) noexcept override {
        return Guarded([&] {
            IDispatch* dispatch = RequireDispatchShape(shape);
            ReleaseDispatch release{dispatch};
            // Shape.Type is read here rather than passed in: this entry point has
            // no classification to reuse, and an unkeyable Shape was never in the
            // record to begin with.
            Value type = get(dispatch, L"Type");
            const office::ShapeKey key = office::DescribeShape(dispatch, type.integer());
            if (key.valid) {
                office::ForgetApplied(key);
            }
        });
    }

    void ClearPictureCache() noexcept override {
        PictureIdentities::Instance().Clear();
        office::ForgetAllApplied();
    }

    void PictureCacheStats(std::size_t* textures,
                           std::size_t* shapes,
                           std::uint64_t* skipped) const noexcept override {
        const office::SkipStats stats = office::ApplySkipStats();
        if (textures) {
            // What the picture path holds here is one identity per path, not a
            // decoded image - see PictureIdentities. The count still answers the
            // question a caller asks of it: how many distinct images are known.
            *textures = PictureIdentities::Instance().Count();
        }
        if (shapes) {
            *shapes = stats.shapes;
        }
        if (skipped) {
            *skipped = stats.skipped;
        }
    }

    std::size_t TextureCount() const noexcept override {
        return portable::TextureCount();
    }
};

} // namespace

std::unique_ptr<Backend> CreateBackend() {
    return std::make_unique<PortableOfficeBackend>();
}

} // namespace bb
