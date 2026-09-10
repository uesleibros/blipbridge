/**
 * @file apply_skip_cache.cpp
 * The per-Shape "already carries this image" record.
 *
 * Deliberately small. All the hard thinking is in what it refuses to remember -
 * see shape_identity.hpp for why groups are excluded and why the key is values
 * rather than pointers - and in the verification below.
 */

#include "apply_skip_cache.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <map>

namespace bb::office {
namespace {

/// `Fill.Type` for a picture fill: what a successful apply leaves behind.
constexpr long kPictureFill = 6;

/**
 * Process-wide, single-apartment, like the texture store it accompanies.
 *
 * Holds nothing but numbers, so it never needs telling that a Shape or a
 * document went away: a stale entry can only ever cause one unnecessary apply,
 * never a wrong one, because the fill is re-verified before any skip.
 */
class SkipCache {
  public:
    static SkipCache& Instance() {
        static SkipCache cache;
        return cache;
    }

    void Remember(const ShapeKey& key, std::uint64_t textureId) {
        if (key.valid) {
            shapes_[key] = textureId;
        }
    }

    bool Matches(const ShapeKey& key, std::uint64_t textureId, IDispatch* shape) {
        if (!key.valid || textureId == 0) {
            return false;
        }
        const auto found = shapes_.find(key);
        if (found == shapes_.end() || found->second != textureId) {
            return false;
        }

        /*
         * The record says this image is already on this Shape. Confirm the fill
         * is still a picture before trusting it.
         *
         * This is the whole safety of the skip: a fill cleared or replaced with a
         * colour must be re-applied, and reading Fill.Type is about three
         * microseconds against the ~170 the skip saves. It cannot detect a fill
         * replaced by a *different* picture - that still reads as 6 - which is
         * why BB_InvalidateShape exists and is documented.
         */
        try {
            bb::Value fill = bb::get(shape, L"Fill");
            if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
                shapes_.erase(found);
                return false;
            }
            if (bb::get(fill.obj(), L"Type").integer() != kPictureFill) {
                shapes_.erase(found);
                return false;
            }
        } catch (...) {
            // A Shape that will not answer is not one to skip work on.
            shapes_.erase(found);
            return false;
        }

        ++skipped_;
        return true;
    }

    void Forget(const ShapeKey& key) {
        if (key.valid) {
            shapes_.erase(key);
        }
    }

    void ForgetAll() noexcept {
        shapes_.clear();
    }

    bool Any() const noexcept {
        return !shapes_.empty();
    }

    void ForgetTexture(std::uint64_t textureId) noexcept {
        for (auto entry = shapes_.begin(); entry != shapes_.end();) {
            entry = entry->second == textureId ? shapes_.erase(entry) : std::next(entry);
        }
    }

    SkipStats Stats() const noexcept {
        return SkipStats{shapes_.size(), skipped_};
    }

  private:
    SkipCache() = default;

    std::map<ShapeKey, std::uint64_t> shapes_;
    std::uint64_t skipped_ = 0;
};

} // namespace

void RememberApplied(const ShapeKey& key, std::uint64_t textureId) noexcept {
    try {
        SkipCache::Instance().Remember(key, textureId);
    } catch (...) {
        // Failing to remember costs one apply next time. It must never fail the
        // apply that just succeeded.
    }
}

bool AlreadyCarries(const ShapeKey& key, std::uint64_t textureId, IDispatch* shape) noexcept {
    try {
        return SkipCache::Instance().Matches(key, textureId, shape);
    } catch (...) {
        return false;
    }
}

void ForgetApplied(const ShapeKey& key) noexcept {
    try {
        SkipCache::Instance().Forget(key);
    } catch (...) {
    }
}

void ForgetAllApplied() noexcept {
    SkipCache::Instance().ForgetAll();
}

void ForgetTexture(std::uint64_t textureId) noexcept {
    SkipCache::Instance().ForgetTexture(textureId);
}

bool AnyRemembered() noexcept {
    return SkipCache::Instance().Any();
}

SkipStats ApplySkipStats() noexcept {
    return SkipCache::Instance().Stats();
}

} // namespace bb::office
