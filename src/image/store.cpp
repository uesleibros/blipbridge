/**
 * @file store.cpp
 * The CPU image store. See store.hpp for why images and textures are separate.
 */

#include "store.hpp"

#include <map>

namespace bb::image {
namespace {

/**
 * Process-wide, single-apartment, like the texture store it sits beside.
 *
 * Handles are never recycled. A released handle stays stale for the life of the
 * process, so a stale handle can only ever be refused - it can never silently
 * resolve to a different image that happens to have taken its number.
 */
class ImageStore {
  public:
    static ImageStore& Instance() {
        static ImageStore store;
        return store;
    }

    std::uint64_t Add(ImageRef image) noexcept {
        if (!image || image->pixels.empty()) {
            return 0;
        }
        try {
            const std::uint64_t handle = next_++;
            images_.emplace(handle, std::move(image));
            return handle;
        } catch (...) {
            return 0;
        }
    }

    ImageRef Find(std::uint64_t handle) const noexcept {
        const auto found = images_.find(handle);
        return found == images_.end() ? ImageRef{} : found->second;
    }

    bool Release(std::uint64_t handle) noexcept {
        return images_.erase(handle) != 0;
    }

    void Clear() noexcept {
        images_.clear();
    }

    std::uint32_t Count() const noexcept {
        return static_cast<std::uint32_t>(images_.size());
    }

  private:
    ImageStore() = default;

    std::map<std::uint64_t, ImageRef> images_;
    std::uint64_t next_ = kImageHandleBase;
};

} // namespace

bool IsImageHandle(std::uint64_t handle) noexcept {
    return handle >= kImageHandleBase;
}

std::uint64_t AddImage(ImageRef image) noexcept {
    return ImageStore::Instance().Add(std::move(image));
}

ImageRef FindImage(std::uint64_t handle) noexcept {
    // The range check comes first so a texture handle is refused as "not an
    // image" rather than searched for and coincidentally missed.
    if (!IsImageHandle(handle)) {
        return ImageRef{};
    }
    return ImageStore::Instance().Find(handle);
}

bool ReleaseImage(std::uint64_t handle) noexcept {
    if (!IsImageHandle(handle)) {
        return false;
    }
    return ImageStore::Instance().Release(handle);
}

void ClearImages() noexcept {
    ImageStore::Instance().Clear();
}

std::uint32_t ImageCount() noexcept {
    return ImageStore::Instance().Count();
}

} // namespace bb::image
