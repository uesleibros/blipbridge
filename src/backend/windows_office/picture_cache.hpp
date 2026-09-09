#pragma once
/**
 * @file picture_cache.hpp
 * The one-call picture fill, and the two caches that make repeating it cheap.
 *
 * This is what `BB_ApplyPicture` and the VBA `UserPicture2` are built on. It
 * makes the decision the caller should not have to: native path where the Shape
 * class has a validated one, ordinary `Fill.UserPicture` where it does not, and
 * a specific refusal where neither applies.
 *
 * Two caches sit underneath, both keyed on things that are cheap to check:
 *
 *  - **path to texture**, so a file is read and decoded once no matter how many
 *    Shapes it is applied to;
 *  - **Shape to last texture**, so applying the same image to the same Shape
 *    twice in a row costs no Office transaction at all.
 *
 * Neither cache holds a Shape, a receiver or any other document object. The
 * Shape cache is keyed by *identity values* read out of the Shape - not by
 * pointer - because a pointer would outlive the object it named. Nothing here
 * has to be told when a Shape dies.
 *
 * A Shape that cannot yield a complete key is not cached at all, which costs one
 * apply and never risks a wrong one.
 *
 * Threading: STA only, like everything else that touches these objects.
 */

#include <windows.h>
#include <oleauto.h>

#include <cstdint>
#include <string>

namespace bb::office {

/**
 * Fills @p shape with the image at @p path, choosing the fastest safe route.
 *
 * Throws bb::Error naming the specific reason on failure - an unvalidated Office
 * build, an unreadable file, a Shape class with no picture-fill path at all -
 * rather than a generic failure. The fallback is used only for Shape classes
 * that have no native path, never to paper over a native failure.
 */
void ApplyPictureCached(IDispatch* shape, const std::wstring& path);

/**
 * Forgets what was last applied to @p shape.
 *
 * Call after changing that Shape's fill by any other means; the next apply then
 * does real work instead of being skipped. Forgetting an unknown Shape is not an
 * error.
 */
void InvalidateShapeCache(IDispatch* shape);

/// Releases every cached texture and forgets every Shape. Safe at any time.
void ClearPictureCache() noexcept;

/// What the caches currently hold, plus the running total of applies avoided.
struct PictureCacheStats {
    std::size_t textures = 0;
    std::size_t shapes = 0;
    std::uint64_t skipped = 0;
};

PictureCacheStats GetPictureCacheStats() noexcept;

} // namespace bb::office
