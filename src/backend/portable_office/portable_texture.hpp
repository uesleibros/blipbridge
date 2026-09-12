#pragma once
/**
 * @file portable_texture.hpp
 * The texture resource for the portable backend: decoded pixels, plus the
 * temporary PNG that lets Office see them.
 *
 * ## How this differs from the accelerated texture, and why
 *
 * `bb::office::NativeTexture` holds two opaque GFX pointers and **no pixels**:
 * the image lives inside Office's own image cache, and applying it is a private
 * call that hands Office a resource it already owns.
 *
 * None of that is reachable without the reverse-engineered layouts, so the
 * portable backend goes the only documented way there is: `Fill.UserPicture`,
 * which takes a **path**. A texture here therefore holds the decoded BGRA
 * pixels, and writes them to a temporary PNG the first time one is applied.
 *
 * That gives the two resources opposite shapes, and it is worth stating plainly
 * rather than discovering later:
 *
 * | | accelerated | portable |
 * |---|---|---|
 * | pixels retained | no | yes, BGRA32 |
 * | cost of the first apply | one private call | encode + write, then Office reads the file |
 * | cost of a repeat apply | one private call | Office reads the same file again |
 * | what Office stores | a shared cached image | an embedded copy per fill |
 *
 * The PNG is written **once per texture, not once per apply**: encoding is the
 * expensive part and the file does not change, so a texture applied to fifty
 * Shapes encodes once. The file lives as long as the texture and is deleted with
 * it.
 *
 * ## Handles
 *
 * Handles come from the same numeric space as the accelerated backend -
 * `0x1000000` upwards - so a handle means the same kind of thing on both, and
 * the ABI's refusal of an *image* handle where a *texture* was wanted behaves
 * identically. Handles are never recycled: a released one stays permanently
 * stale rather than becoming somebody else's texture.
 *
 * ## Ownership
 *
 * Two owners are possible, exactly as on the accelerated backend:
 *
 *     public handle   -> TextureRef
 *     picture cache   -> TextureRef
 *
 * Dropping either leaves the other working.
 *
 * Threading: single-threaded-apartment, like everything else that touches these
 * objects.
 */

#include <cstdint>
#include <memory>
#include <string>
#include <vector>

namespace bb::portable {

/// One decoded image, and the temporary file Office reads it from.
class PortableTexture {
  public:
    PortableTexture(std::vector<std::uint8_t> pixels,
                    std::uint32_t width,
                    std::uint32_t height,
                    std::uint64_t id);
    ~PortableTexture();

    PortableTexture(const PortableTexture&) = delete;
    PortableTexture& operator=(const PortableTexture&) = delete;

    /**
     * The path to a PNG holding these pixels, encoded and written on first call.
     *
     * Throws bb::Error if the image cannot be encoded or the file cannot be
     * written - both of which are real failures to report, never something to
     * paper over, because the apply that needed the file cannot happen without
     * it.
     */
    const std::wstring& Path();

    std::uint64_t id() const noexcept {
        return id_;
    }
    std::uint32_t width() const noexcept {
        return width_;
    }
    std::uint32_t height() const noexcept {
        return height_;
    }
    /// Decoded bytes held. Diagnostics only.
    std::size_t byteCount() const noexcept {
        return pixels_.size();
    }
    /// Whether the temporary PNG has been written yet. Diagnostics only.
    bool encoded() const noexcept {
        return !path_.empty();
    }
    /// How many applies this image has served. Diagnostics only.
    unsigned long applyCount() const noexcept {
        return applyCount_;
    }
    void CountApplies(unsigned long times) noexcept {
        applyCount_ += times;
    }

  private:
    std::vector<std::uint8_t> pixels_;
    std::uint32_t width_ = 0;
    std::uint32_t height_ = 0;
    std::uint64_t id_ = 0;
    std::wstring path_;
    unsigned long applyCount_ = 0;
};

using TextureRef = std::shared_ptr<PortableTexture>;

/**
 * The next process-unique image id, from the same counter textures use.
 *
 * Exposed because the portable backend also gives an *identity* to each image
 * path it applies, and those identities share one skip-cache record with texture
 * identities. Two counters would eventually produce the same number for a path
 * and a texture, and the skip cache would then believe a Shape carrying one was
 * carrying the other - a wrong picture, silently. One counter makes that
 * impossible rather than unlikely.
 */
std::uint64_t NextSharedImageId() noexcept;

/**
 * How many times this process has actually changed a fill through Office.
 *
 * The portable analogue of the accelerated backend's private-apply entry
 * counter, and it exists for the same reason: a skip has to be provable to have
 * *not happened*, not merely to have been survived. Comparing this across a call
 * says whether Office was entered at all, which is the only thing that makes a
 * skip worth having.
 *
 * Counts fills, not applies: one range fill is one Office edit however many
 * Shapes it covers, which is exactly what the number is asked about.
 */
std::uint64_t OfficeFillCount() noexcept;

/// Records one fill. Called by the backend at the single place it edits a document.
void CountOfficeFill() noexcept;

/// Decodes encoded image bytes into a new texture. The caller owns the only reference.
TextureRef CreateTextureFromBytes(const std::uint8_t* bytes, std::size_t length);

/// Same, from a raw BGRA32 buffer, which is copied.
TextureRef CreateTextureFromPixels(const void* pixels,
                                   std::uint32_t width,
                                   std::uint32_t height,
                                   std::int32_t stride);

/**
 * Gives @p texture a public handle, adding one owner.
 *
 * Handles are never recycled, so a handle released later stays permanently
 * stale even though the image behind it may still be alive elsewhere.
 */
std::uint64_t RegisterTexture(TextureRef texture);

/// Resolves a public handle. Throws bb::Error if it is unknown or was released.
TextureRef LookupTexture(std::uint64_t handle);

/**
 * True when @p handle is a value from this store's handle space at all.
 *
 * Distinct from OwnsHandle, and the distinction matters for diagnosis rather
 * than for correctness. A caller who routes on "do I own this" sends a handle
 * that was never issued somewhere else entirely, and that somewhere else answers
 * with its own vaguer message - so the caller is told "not found" when the store
 * could have told them the handle never existed. Routing on the space and
 * diagnosing in the store keeps the specific answer.
 */
bool IsTextureHandle(std::uint64_t handle) noexcept;

/**
 * True when @p handle names a live texture in this store.
 *
 * The question a caller asks before deciding whose handle this is, so it answers
 * rather than throws: an unknown handle is an ordinary "no", not a failure.
 */
bool OwnsHandle(std::uint64_t handle) noexcept;

/**
 * A process-unique id for one decoded image.
 *
 * The skip cache remembers which image a Shape last received, and must compare
 * that without keeping the image alive or dereferencing something freed.
 */
std::uint64_t TextureIdOf(const TextureRef& texture);

/// Drops the public handle's reference. Throws bb::Error if it is unknown.
void ReleaseTexture(std::uint64_t handle);

/// Drops every public handle's reference. Safe at any time.
void ClearTextures() noexcept;

/// How many public handles are live.
std::size_t TextureCount() noexcept;

/**
 * Diagnostics for one handle, or for the store as a whole when @p handle is 0.
 *
 * A `key=value;` string, in the same shape the accelerated store reports, so the
 * Office harnesses can read either backend's answer with one parser. The fields
 * that only make sense here - whether a temporary file has been written yet -
 * are additions rather than replacements.
 *
 * Never throws: a handle that is gone is reported as gone. This is a diagnostic,
 * and one that failed when asked about a released handle would be useless in
 * exactly the situation it is wanted.
 */
std::wstring DescribeTextures(std::uint64_t handle) noexcept;

/**
 * Deletes the process's temporary directory if it is empty.
 *
 * Called at shutdown. Every texture deletes its own file, so this only removes
 * the directory they were in - and only when nothing is left in it, so a file
 * that could not be deleted is never taken down with something else.
 */
void CleanUpTemporaryDirectory() noexcept;

} // namespace bb::portable
