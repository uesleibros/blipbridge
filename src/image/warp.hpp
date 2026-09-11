#pragma once
/**
 * @file warp.hpp
 * Projective (perspective) mapping of an image onto a caller-supplied quad.
 *
 * ## Why this can work at all
 *
 * `tools/probe_freeform_uv_mapping.ps1` answered the question this rests on.
 * PowerPoint maps a picture fill **linearly onto the Shape's bounding box and
 * then clips it to the path** - measured across a rectangle, a parallelogram, a
 * trapezoid, an extreme trapezoid, a rotated quad and a perspective quad, with a
 * worst error of 0.8 source texels, which is the quantisation floor of the probe
 * itself. The four nodes of a Freeform decide what is *visible*. They do not
 * move the texture.
 *
 * That is what makes pre-warping honest rather than hopeful. If the image is
 * warped into the quad's own bounding box here, then Office's mapping of that
 * image onto the same bounding box is the identity, and the path clips away
 * everything outside the quad. The perspective the caller asked for is the
 * perspective they get, because we did it before Office saw the image.
 *
 * ## What is compensated, and what is not
 *
 * Compensated: where each source texel lands inside the quad. That is entirely
 * this file's arithmetic.
 *
 * Not compensated, and not ours: how Office then draws the resulting image when
 * the slide is zoomed, scaled, printed or shown. Office's own sampler runs
 * downstream of everything here, and no amount of pre-warping reaches it. The
 * supported claim is about the image Office receives.
 *
 * ## Point order, fixed
 *
 *     p[0] = source top-left      (u,v) = (0,0)
 *     p[1] = source top-right             (1,0)
 *     p[2] = source bottom-right          (1,1)
 *     p[3] = source bottom-left           (0,1)
 *
 * The points are never reordered to suit the screen. A caller who hands them in
 * a different order is asking for a flipped or crossed mapping and will get one,
 * which is the only behaviour that lets a caller mirror a face on purpose.
 *
 * ## Coordinates
 *
 * Points are in whatever unit the caller works in - PowerPoint points, pixels,
 * anything - because only their *relative* geometry matters. The output is
 * rasterised into the quad's bounding box, and `WarpResult` reports that box so
 * the caller can position the Shape to match it.
 */

#include "decode.hpp"
#include "resample.hpp"

#include <cstdint>

namespace bb::image {

/// One corner, in the caller's own coordinate space.
struct PointF {
    float x = 0.0f;
    float y = 0.0f;
};

/// Why a warp was refused. Each says something specific about the quad.
enum class WarpStatus {
    Ok,
    NullSource,
    InvalidSourceSize,
    /// A coordinate was infinite or not-a-number.
    NonFinitePoint,
    /// Two or more corners are the same point.
    DuplicatePoints,
    /// The quad encloses no area, or is so thin it has no interior.
    DegenerateQuad,
    /// The quad's bounding box is larger than the resampler's pixel ceiling.
    TooLarge,
    /// The projective transform could not be inverted - a collinear quad.
    NotInvertible,
    UnknownFilter,
    OutOfMemory,
};

const char* DescribeWarpStatus(WarpStatus status) noexcept;

/// What a warp produced, and where the caller should put it.
struct WarpResult {
    WarpStatus status = WarpStatus::Ok;
    DecodedImage image;

    /// The quad's bounding box in the caller's coordinates. The image covers
    /// exactly this rectangle, so a Shape placed here shows the quad unmoved.
    float left = 0.0f;
    float top = 0.0f;
    float width = 0.0f;
    float height = 0.0f;

    bool ok() const noexcept {
        return status == WarpStatus::Ok;
    }
};

/**
 * Warps @p source onto @p quad, rasterised into the quad's bounding box.
 *
 * @param source        BGRA32, @p sourceHeight rows of @p sourceStride bytes.
 * @param quad          four corners, in the documented order.
 * @param outputWidth   raster width in pixels; 0 for one pixel per unit of the
 *                      bounding box, which is the caller's own scale.
 * @param outputHeight  raster height in pixels; 0 for the same.
 * @param filter        Nearest or Bilinear. See below about Bicubic.
 *
 * Pixels outside the quad are left fully transparent, so the result is correct
 * whether or not the Shape's own path also clips them - a caller filling a
 * rectangular Shape with a trapezoid gets the trapezoid.
 *
 * The mapping is a true projective transform, not an affine approximation: a
 * trapezoid gets foreshortening that varies across it, which is the entire point
 * of the exercise. Sampling is per-pixel inverse-mapped through the homography,
 * so the source never has to be subdivided.
 *
 * Nearest is exact point sampling and is what pixel art needs. Bilinear
 * interpolates in premultiplied alpha, for the same reason `resample.hpp` does.
 * **Bicubic is deliberately not offered here**: in a projective sampler it costs
 * sixteen taps per output pixel with a varying kernel footprint, and the
 * measured benefit over bilinear on the images this is for did not come close to
 * justifying that in a per-frame path. Adding it for API symmetry would be
 * adding a cost nobody asked for.
 */
WarpResult WarpQuad(const std::uint8_t* source,
                    std::uint32_t sourceWidth,
                    std::uint32_t sourceHeight,
                    std::int32_t sourceStride,
                    const PointF quad[4],
                    std::uint32_t outputWidth,
                    std::uint32_t outputHeight,
                    ScaleFilter filter) noexcept;

} // namespace bb::image
