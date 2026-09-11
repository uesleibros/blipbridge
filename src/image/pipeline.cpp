/**
 * @file pipeline.cpp
 * Crop, orient, resize. See pipeline.hpp for the stage order and why it is that.
 */

#include "pipeline.hpp"

#include <cstring>
#include <new>

namespace bb::image {
namespace {

/// Copies a rectangle out of a strided source into a tightly packed buffer.
bool CropInto(const std::uint8_t* source,
              std::int32_t sourceStride,
              std::uint32_t x,
              std::uint32_t y,
              std::uint32_t width,
              std::uint32_t height,
              DecodedImage& out) noexcept {
    try {
        out.pixels.resize(static_cast<std::size_t>(width) * height * 4);
    } catch (const std::bad_alloc&) {
        return false;
    }
    out.width = width;
    out.height = height;

    const auto stride = static_cast<std::size_t>(sourceStride);
    const std::size_t rowBytes = static_cast<std::size_t>(width) * 4;
    for (std::uint32_t row = 0; row < height; ++row) {
        const std::uint8_t* from =
            source + (static_cast<std::size_t>(y) + row) * stride + static_cast<std::size_t>(x) * 4;
        std::memcpy(&out.pixels[static_cast<std::size_t>(row) * rowBytes], from, rowBytes);
    }
    return true;
}

/**
 * Applies an orientation change into a fresh buffer.
 *
 * Every output pixel is exactly one input pixel - no averaging, no rounding -
 * so these are lossless whatever the filter is and pixel-art survives them.
 */
bool TransformInto(const DecodedImage& source, Transform transform, DecodedImage& out) noexcept {
    const std::uint32_t width = source.width;
    const std::uint32_t height = source.height;
    const bool swaps = transform == Transform::Rotate90 || transform == Transform::Rotate270;

    out.width = swaps ? height : width;
    out.height = swaps ? width : height;
    try {
        out.pixels.resize(static_cast<std::size_t>(width) * height * 4);
    } catch (const std::bad_alloc&) {
        return false;
    }

    const auto* in = source.pixels.data();
    auto* result = out.pixels.data();
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            std::uint32_t targetX = x;
            std::uint32_t targetY = y;
            switch (transform) {
            case Transform::None:
                break;
            case Transform::FlipHorizontal:
                targetX = width - 1 - x;
                break;
            case Transform::FlipVertical:
                targetY = height - 1 - y;
                break;
            case Transform::Rotate90:
                // clockwise: the top row becomes the right column
                targetX = height - 1 - y;
                targetY = x;
                break;
            case Transform::Rotate180:
                targetX = width - 1 - x;
                targetY = height - 1 - y;
                break;
            case Transform::Rotate270:
                targetX = y;
                targetY = width - 1 - x;
                break;
            }
            const std::size_t from = (static_cast<std::size_t>(y) * width + x) * 4;
            const std::size_t to = (static_cast<std::size_t>(targetY) * out.width + targetX) * 4;
            std::memcpy(result + to, in + from, 4);
        }
    }
    return true;
}

} // namespace

bool IsSupportedTransform(std::uint32_t transform) noexcept {
    return transform <= static_cast<std::uint32_t>(Transform::Rotate270);
}

const char* TransformName(Transform transform) noexcept {
    switch (transform) {
    case Transform::None:
        return "none";
    case Transform::FlipHorizontal:
        return "flip-horizontal";
    case Transform::FlipVertical:
        return "flip-vertical";
    case Transform::Rotate90:
        return "rotate-90";
    case Transform::Rotate180:
        return "rotate-180";
    case Transform::Rotate270:
        return "rotate-270";
    }
    return "unknown";
}

const char* DescribePrepareStatus(PrepareStatus status) noexcept {
    switch (status) {
    case PrepareStatus::Ok:
        return "The image was prepared";
    case PrepareStatus::InvalidCrop:
        return "The region is empty or reaches outside the image";
    case PrepareStatus::UnknownTransform:
        return "That transform is not one this build implements";
    case PrepareStatus::DecodeFailed:
        return "The image could not be decoded";
    case PrepareStatus::ResampleFailed:
        return "The image could not be resampled";
    case PrepareStatus::OutOfMemory:
        return "Not enough memory to prepare the image";
    }
    return "Unknown prepare status";
}

bool PrepareRequest::isIdentity(std::uint32_t sourceWidth,
                                std::uint32_t sourceHeight) const noexcept {
    const bool wholeImage = (cropWidth == 0 || cropHeight == 0) ||
                            (cropX == 0 && cropY == 0 && cropWidth == sourceWidth &&
                             cropHeight == sourceHeight);
    const bool sameSize = (targetWidth == 0 || targetHeight == 0) ||
                          (targetWidth == sourceWidth && targetHeight == sourceHeight);
    return wholeImage && sameSize && transform == Transform::None;
}

PrepareResult PreparePixels(const std::uint8_t* source,
                            std::uint32_t sourceWidth,
                            std::uint32_t sourceHeight,
                            std::int32_t sourceStride,
                            const PrepareRequest& request) noexcept {
    PrepareResult result;

    if (!source || sourceWidth == 0 || sourceHeight == 0) {
        result.status = PrepareStatus::DecodeFailed;
        result.decode = DecodeStatus::NoData;
        return result;
    }
    if (sourceStride < static_cast<std::int32_t>(sourceWidth) * 4) {
        result.status = PrepareStatus::DecodeFailed;
        result.decode = DecodeStatus::UnsupportedSize;
        return result;
    }
    if (!IsSupportedTransform(static_cast<std::uint32_t>(request.transform))) {
        result.status = PrepareStatus::UnknownTransform;
        return result;
    }

    /*
     * The crop rectangle, resolved and checked in 64-bit.
     *
     * A zero width or height means the whole image rather than an error,
     * because that is what a default-constructed request means and a caller who
     * only wants to resize should not have to name the source size to do it.
     * Anything else must fit: x + width has to stay inside the source, computed
     * wide so that a near-maximal x cannot wrap past it.
     */
    std::uint32_t cropWidth = request.cropWidth;
    std::uint32_t cropHeight = request.cropHeight;
    if (cropWidth == 0 || cropHeight == 0) {
        if (request.cropX != 0 || request.cropY != 0) {
            // An origin with no size is a caller who meant something specific
            // and got it wrong, rather than a caller who omitted the crop.
            result.status = PrepareStatus::InvalidCrop;
            return result;
        }
        cropWidth = sourceWidth;
        cropHeight = sourceHeight;
    }
    const std::uint64_t right = static_cast<std::uint64_t>(request.cropX) + cropWidth;
    const std::uint64_t bottom = static_cast<std::uint64_t>(request.cropY) + cropHeight;
    if (right > sourceWidth || bottom > sourceHeight) {
        result.status = PrepareStatus::InvalidCrop;
        return result;
    }

    DecodedImage cropped;
    if (!CropInto(source,
                  sourceStride,
                  request.cropX,
                  request.cropY,
                  cropWidth,
                  cropHeight,
                  cropped)) {
        result.status = PrepareStatus::OutOfMemory;
        return result;
    }

    if (request.transform != Transform::None) {
        DecodedImage oriented;
        if (!TransformInto(cropped, request.transform, oriented)) {
            result.status = PrepareStatus::OutOfMemory;
            return result;
        }
        cropped = std::move(oriented);
    }

    const std::uint32_t targetWidth =
        request.targetWidth == 0 ? cropped.width : request.targetWidth;
    const std::uint32_t targetHeight =
        request.targetHeight == 0 ? cropped.height : request.targetHeight;

    if (targetWidth == cropped.width && targetHeight == cropped.height) {
        // Already the right size. Resample would copy the rows and return them
        // bit-exact anyway; skipping it says so plainly and saves the copy.
        result.image = std::move(cropped);
        return result;
    }

    std::vector<std::uint8_t> scaled;
    const ResampleStatus status = Resample(cropped.pixels.data(),
                                           cropped.width,
                                           cropped.height,
                                           cropped.stride(),
                                           targetWidth,
                                           targetHeight,
                                           request.filter,
                                           scaled);
    if (status != ResampleStatus::Ok) {
        result.status = PrepareStatus::ResampleFailed;
        result.resample = status;
        return result;
    }

    result.image.width = targetWidth;
    result.image.height = targetHeight;
    result.image.pixels = std::move(scaled);
    return result;
}

PrepareResult PrepareEncoded(const std::uint8_t* bytes,
                             std::size_t length,
                             const PrepareRequest& request) noexcept {
    PrepareResult result;
    DecodedImage decoded;
    const DecodeStatus status = Decode(bytes, length, decoded);
    if (status != DecodeStatus::Ok) {
        result.status = PrepareStatus::DecodeFailed;
        result.decode = status;
        return result;
    }
    return PreparePixels(
        decoded.pixels.data(), decoded.width, decoded.height, decoded.stride(), request);
}

PrepareResult PrepareFile(const wchar_t* path, const PrepareRequest& request) noexcept {
    PrepareResult result;
    DecodedImage decoded;
    const DecodeStatus status = DecodeFile(path, decoded);
    if (status != DecodeStatus::Ok) {
        result.status = PrepareStatus::DecodeFailed;
        result.decode = status;
        return result;
    }
    return PreparePixels(
        decoded.pixels.data(), decoded.width, decoded.height, decoded.stride(), request);
}

} // namespace bb::image
