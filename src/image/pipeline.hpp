#pragma once
/**
 * @file pipeline.hpp
 * Preparing an image before Office ever sees it: crop, orient, resize.
 *
 * ## Why one function and not six
 *
 * "Crop", "scale" and "flip" are not three features. They are three stages of
 * the same journey from an encoded file to the pixels a Shape will show, and a
 * caller pulling one tile out of an atlas usually wants two of them at once.
 * Six entry points would mean six places to validate sizes, six chances to
 * disagree about pixel centres, and a caller doing `LoadRegion` then
 * `LoadScaled` would decode the file twice.
 *
 * So there is one request, with stages that switch themselves off when the
 * caller does not ask for them.
 *
 * ## The order, fixed and documented
 *
 *     decode  ->  crop  ->  transform  ->  resize
 *
 * Crop first, because a region is expressed in the source image's own
 * coordinates and that is the only frame a caller can name it in. Transform
 * before resize, so that `targetWidth` and `targetHeight` always mean the size
 * of the image that comes out - rotating by 90 after resizing would make the
 * target size mean the size of something the caller never sees.
 *
 * ## What this deliberately does not do
 *
 * It does not touch Office, it does not create a texture, and it does not read a
 * Shape. It turns bytes into other bytes. The texture-creating layer above
 * decides what to do with the result, which is what keeps this file testable
 * without PowerPoint and usable on x86.
 *
 * No filter maths lives here either: `resample.hpp` remains the one place any of
 * that exists, and this calls it.
 */

#include "decode.hpp"
#include "resample.hpp"

#include <cstdint>
#include <vector>

namespace bb::image {

/**
 * An orientation change, applied after cropping and before resizing.
 *
 * These are the lossless ones - every output pixel is exactly some input pixel,
 * so they are exact whatever the filter is, and they are cheap. Arbitrary-angle
 * rotation is not here on purpose: it needs resampling, a background colour and
 * a decision about the output size, and none of those are free choices.
 */
enum class Transform : std::uint32_t {
    None = 0,
    FlipHorizontal = 1,
    FlipVertical = 2,
    Rotate90 = 3,  ///< clockwise
    Rotate180 = 4,
    Rotate270 = 5, ///< clockwise, i.e. 90 anticlockwise
};

/// True for a value the implementation actually supports.
bool IsSupportedTransform(std::uint32_t transform) noexcept;

/// Human-readable name, for diagnostics and error text.
const char* TransformName(Transform transform) noexcept;

/**
 * What to do to an image on the way in. Every field has an "unset" value that
 * switches its stage off, so the common case stays short.
 */
struct PrepareRequest {
    /// Region of the source, in source pixels. A zero width or height means the
    /// whole image, so a default-constructed request crops nothing.
    std::uint32_t cropX = 0;
    std::uint32_t cropY = 0;
    std::uint32_t cropWidth = 0;
    std::uint32_t cropHeight = 0;

    Transform transform = Transform::None;

    /// Size of the result. Zero for either means "whatever the previous stages
    /// produced", so a request can crop without resizing or resize without
    /// cropping.
    std::uint32_t targetWidth = 0;
    std::uint32_t targetHeight = 0;

    ScaleFilter filter = ScaleFilter::Bilinear;

    /// True when nothing would change the pixels at all.
    bool isIdentity(std::uint32_t sourceWidth, std::uint32_t sourceHeight) const noexcept;
};

/// Why a prepare was refused. Decode and resample reasons pass through as-is.
enum class PrepareStatus {
    Ok,
    /// The crop rectangle is empty, or reaches outside the source image.
    InvalidCrop,
    UnknownTransform,
    /// A decode failed; see the DecodeStatus the call also reports.
    DecodeFailed,
    /// A resample failed; see the ResampleStatus the call also reports.
    ResampleFailed,
    OutOfMemory,
};

const char* DescribePrepareStatus(PrepareStatus status) noexcept;

/// What a prepare produced, and what each stage said if it did not.
struct PrepareResult {
    PrepareStatus status = PrepareStatus::Ok;
    DecodeStatus decode = DecodeStatus::Ok;
    ResampleStatus resample = ResampleStatus::Ok;
    DecodedImage image;

    bool ok() const noexcept {
        return status == PrepareStatus::Ok;
    }
};

/**
 * Runs the stages over already-decoded BGRA32.
 *
 * @param source  BGRA32, @p sourceHeight rows of @p sourceStride bytes.
 *
 * Every rectangle is checked against the source in 64-bit before a byte is read,
 * so a crop that reaches outside the image is a refusal rather than a read past
 * the end. A request that changes nothing copies the rows and returns them,
 * which keeps "no transform" honest rather than accidentally exact.
 */
PrepareResult PreparePixels(const std::uint8_t* source,
                            std::uint32_t sourceWidth,
                            std::uint32_t sourceHeight,
                            std::int32_t sourceStride,
                            const PrepareRequest& request) noexcept;

/// Decodes @p bytes, then runs the stages. One decode, whatever the request asks.
PrepareResult PrepareEncoded(const std::uint8_t* bytes,
                             std::size_t length,
                             const PrepareRequest& request) noexcept;

/// Reads and decodes @p path, then runs the stages.
PrepareResult PrepareFile(const wchar_t* path, const PrepareRequest& request) noexcept;

} // namespace bb::image
