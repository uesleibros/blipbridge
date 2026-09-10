#pragma once
/**
 * @file resample.hpp
 * Software resampling for raw BGRA32 images.
 *
 * This is ordinary, portable image code: no Office, no architecture assumptions,
 * no undocumented anything. It runs on x86 and x64 alike and is unit-testable
 * without PowerPoint, which is why it lives under `src/image/` rather than in a
 * backend.
 *
 * ## What this does and does not control
 *
 * BlipBridge resamples the pixels **before** handing them to Office:
 *
 *     source BGRA -> BlipBridge filter -> image resource -> Office
 *
 * It says nothing about how Office then draws that image - Shape scaling,
 * slideshow scaling, zoom, DPI and the renderer's own sampling are all
 * downstream and unaffected. The supported claim is only that the image Office
 * receives was resampled by the filter you chose.
 *
 * ## Conventions, chosen once and applied to every filter
 *
 * **Pixel centres.** A destination pixel `d` maps to the source coordinate
 * `(d + 0.5) * scale - 0.5`, so both images are treated as grids of unit squares
 * whose samples sit at their centres. This is what keeps an image from drifting
 * half a pixel when it is scaled, and it is the same convention every filter
 * here uses.
 *
 * **Edges.** Coordinates are clamped to the nearest valid source pixel. No
 * wrapping, no mirroring, no invented border colour.
 *
 * **Alpha.** Interpolating filters work in premultiplied alpha and convert back
 * afterwards. Interpolating straight BGRA would blend the colour of fully
 * transparent pixels into visible ones, which is what produces dark halos around
 * cut-out edges. Nearest does no blending, so it is exact either way and is left
 * alone.
 */

#include <cstddef>
#include <cstdint>
#include <vector>

namespace bb::image {

/**
 * How to resample. Values are part of the public C ABI - see `BB_SCALE_*` in
 * blipbridge.h - so they are fixed and never renumbered.
 *
 * Every value listed here is implemented and tested. A filter is not named until
 * it works.
 */
enum class ScaleFilter : std::uint32_t {
    /// Exact point sampling. The chosen source pixel's BGRA is copied unchanged.
    Nearest = 0,
    /// 2x2 interpolation between the four source pixels around the sample point.
    Bilinear = 1,
    /// Catmull-Rom cubic over a 4x4 neighbourhood; see Resample for why that one.
    Bicubic = 2,
};

/// True for a value the implementation actually supports.
bool IsSupportedFilter(std::uint32_t filter) noexcept;

/// Human-readable name, for diagnostics and error text.
const char* FilterName(ScaleFilter filter) noexcept;

/**
 * Why a resample was refused. Callers map these onto their own error space; the
 * point is that each says something different and specific.
 */
enum class ResampleStatus {
    Ok,
    NullSource,
    InvalidSourceSize,
    InvalidTargetSize,
    InvalidStride,
    UnknownFilter,
    TooLarge,
    OutOfMemory,
};

/// A sentence explaining a status, suitable for a caller's error message.
const char* DescribeStatus(ResampleStatus status) noexcept;

/**
 * The largest image this will produce, in pixels.
 *
 * A limit exists so that a bad width and height cannot ask for an allocation
 * that would fail in a less controlled way. 64 megapixels is far above anything
 * a slide needs and still only 256 MB of BGRA.
 */
inline constexpr std::uint64_t kMaxPixels = 64ull * 1024ull * 1024ull;

/**
 * Resamples @p source into a tightly packed BGRA32 buffer of the target size.
 *
 * @param source        BGRA32 pixels, `sourceHeight` rows of `sourceStride` bytes.
 * @param sourceWidth   source pixels per row, > 0.
 * @param sourceHeight  source rows, > 0.
 * @param sourceStride  bytes per source row, at least `sourceWidth * 4`.
 * @param targetWidth   destination pixels per row, > 0.
 * @param targetHeight  destination rows, > 0.
 * @param filter        which filter; must be a supported value.
 * @param out           receives the result, `targetWidth * targetHeight * 4` bytes.
 *
 * Returns `ResampleStatus::Ok` on success. Every size and stride is validated
 * before anything is allocated or read, and every multiplication is done in
 * 64-bit and range-checked, so no combination of arguments can overflow into a
 * short allocation and a long read.
 *
 * When the source and target sizes match, the rows are copied and no filter runs
 * at all - which also means the result is bit-exact, whichever filter was asked
 * for.
 */
ResampleStatus Resample(const std::uint8_t* source,
                        std::uint32_t sourceWidth,
                        std::uint32_t sourceHeight,
                        std::int32_t sourceStride,
                        std::uint32_t targetWidth,
                        std::uint32_t targetHeight,
                        ScaleFilter filter,
                        std::vector<std::uint8_t>& out);

} // namespace bb::image
