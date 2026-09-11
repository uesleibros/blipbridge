#pragma once
/**
 * @file store.hpp
 * The CPU-side image resource: decoded pixels a caller can process repeatedly.
 *
 * ## Why this exists, and why it is not a texture
 *
 * A `BlipBridgeTexture` is an **Office resource**. It holds two opaque GFX
 * pointers and no pixels, because that is what Office hands back and all Office
 * needs. Nothing in a texture handle can be cropped, flipped or warped, and
 * making textures secretly retain a decoded copy would double the memory of
 * every texture in the process to serve the few that get processed twice.
 *
 * So processing gets its own resource. A `BB_Image` is **decoded BGRA32 owned by
 * us**, on the CPU, with no Office in it at all. It exists so that
 *
 *     decode once  ->  warp many times
 *
 * is possible without either resource pretending to be the other. A caller who
 * only wants to put a picture on a Shape never creates one; a caller animating a
 * textured quad every frame creates one and keeps it.
 *
 * The two are converted explicitly, never implicitly: an image becomes a texture
 * when someone asks for a texture. Nothing here ever touches Office, which is
 * also why this file lives under `src/image/` and builds on x86.
 *
 * ## Handles
 *
 * Image handles come from their own numbering space, far from the texture store's,
 * so a texture handle passed to an image call is refused rather than resolving to
 * an unrelated resource. Neither space recycles: a released handle stays stale for
 * the life of the process, which is the same rule the texture store already keeps.
 */

#include "pipeline.hpp"

#include <cstdint>
#include <memory>

namespace bb::image {

/// One decoded image, shared so a handle and any future owner can both hold it.
using ImageRef = std::shared_ptr<const DecodedImage>;

/**
 * Image handles start here, well clear of the texture store's 0x1000000.
 *
 * The gap is the point: handing a texture handle to an image call, or the
 * reverse, must fail rather than find something. Both stores validate their own
 * range before looking anything up.
 */
inline constexpr std::uint64_t kImageHandleBase = 0x4000000ull;

/// True when @p handle is in the image space at all. Says nothing about liveness.
bool IsImageHandle(std::uint64_t handle) noexcept;

/**
 * Registers @p image and returns its handle.
 *
 * Returns 0 for a null or empty image, which is never a valid handle.
 */
std::uint64_t AddImage(ImageRef image) noexcept;

/// The image behind @p handle, or nullptr if it is unknown, stale or not an image.
ImageRef FindImage(std::uint64_t handle) noexcept;

/// Forgets @p handle. False when it was not a live image handle.
bool ReleaseImage(std::uint64_t handle) noexcept;

/// Forgets every image. Handles stay stale rather than being reissued.
void ClearImages() noexcept;

/// How many images are currently held.
std::uint32_t ImageCount() noexcept;

} // namespace bb::image
