/**
 * @file resample.cpp
 * The resampling filters.
 *
 * Three, all implemented and all tested. The conventions they share - pixel
 * centres, edge clamping, premultiplied alpha for the interpolating ones - are
 * documented in resample.hpp and applied identically here, because filters that
 * each invent their own would produce images that shift relative to one another.
 */

#include "resample.hpp"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstring>
#include <new>
#include <utility>

namespace bb::image {
namespace {

constexpr std::uint32_t kChannels = 4;   // B, G, R, A

/// Index of the alpha byte within a BGRA pixel.
constexpr std::size_t kAlpha = 3;

/// Clamps a source coordinate to a real pixel. No wrapping, no border colour.
inline std::int64_t ClampCoordinate(std::int64_t value, std::uint32_t limit) noexcept {
    if (value < 0) {
        return 0;
    }
    const std::int64_t last = static_cast<std::int64_t>(limit) - 1;
    return value > last ? last : value;
}

inline const std::uint8_t* PixelAt(const std::uint8_t* source,
                                   std::size_t stride,
                                   std::int64_t x,
                                   std::int64_t y) noexcept {
    return source + static_cast<std::size_t>(y) * stride + static_cast<std::size_t>(x) * kChannels;
}

/**
 * One source pixel in premultiplied form.
 *
 * Interpolating straight BGRA blends the colour of fully transparent pixels into
 * visible neighbours, which is exactly the dark-halo artefact around cut-out
 * edges. Weighting colour by alpha first makes a transparent pixel contribute
 * nothing but its (zero) alpha, which is the correct behaviour and the reason
 * this type exists.
 */
struct Premultiplied {
    float b = 0.0f;
    float g = 0.0f;
    float r = 0.0f;
    float a = 0.0f;
};

inline Premultiplied LoadPremultiplied(const std::uint8_t* pixel) noexcept {
    const float alpha = static_cast<float>(pixel[kAlpha]) / 255.0f;
    return Premultiplied{static_cast<float>(pixel[0]) * alpha,
                         static_cast<float>(pixel[1]) * alpha,
                         static_cast<float>(pixel[2]) * alpha,
                         static_cast<float>(pixel[kAlpha])};
}

inline std::uint8_t ToByte(float value) noexcept {
    // Cubic filters overshoot by design, so clamping here is required rather
    // than defensive; std::lround gives round-half-away-from-zero.
    const long rounded = std::lround(value);
    if (rounded <= 0) {
        return 0;
    }
    if (rounded >= 255) {
        return 255;
    }
    return static_cast<std::uint8_t>(rounded);
}

/// Converts an accumulated premultiplied sample back to straight BGRA.
inline void StorePremultiplied(const Premultiplied& sum, std::uint8_t* destination) noexcept {
    const float alpha = std::clamp(sum.a, 0.0f, 255.0f);
    if (alpha <= 0.0f) {
        // Fully transparent: colour is meaningless, and writing zeroes keeps the
        // result deterministic rather than leaving whatever the maths produced.
        destination[0] = 0;
        destination[1] = 0;
        destination[2] = 0;
        destination[kAlpha] = 0;
        return;
    }
    const float scale = 255.0f / alpha;
    destination[0] = ToByte(sum.b * scale);
    destination[1] = ToByte(sum.g * scale);
    destination[2] = ToByte(sum.r * scale);
    destination[kAlpha] = ToByte(alpha);
}

/**
 * The source coordinate a destination index samples from.
 *
 * `(d + 0.5) * scale - 0.5` treats both images as grids of unit squares sampled
 * at their centres, which is what stops a scaled image drifting half a pixel.
 */
inline double SourceCoordinate(std::uint32_t destination, double scale) noexcept {
    return (static_cast<double>(destination) + 0.5) * scale - 0.5;
}

void ResampleNearest(const std::uint8_t* source,
                     std::uint32_t sourceWidth,
                     std::uint32_t sourceHeight,
                     std::size_t sourceStride,
                     std::uint32_t targetWidth,
                     std::uint32_t targetHeight,
                     double scaleX,
                     double scaleY,
                     std::uint8_t* out) noexcept {
    for (std::uint32_t y = 0; y < targetHeight; ++y) {
        // Round-half-up on the pixel-centre coordinate. For an integer upscale
        // this reproduces exact block replication: 2x turns AB/CD into
        // AABB/AABB/CCDD/CCDD.
        const std::int64_t sy =
            ClampCoordinate(static_cast<std::int64_t>(std::floor(SourceCoordinate(y, scaleY) + 0.5)),
                            sourceHeight);
        std::uint8_t* row = out + static_cast<std::size_t>(y) * targetWidth * kChannels;
        for (std::uint32_t x = 0; x < targetWidth; ++x) {
            const std::int64_t sx = ClampCoordinate(
                static_cast<std::int64_t>(std::floor(SourceCoordinate(x, scaleX) + 0.5)),
                sourceWidth);
            // Copied whole: no averaging, so the source BGRA survives exactly,
            // alpha included.
            std::memcpy(row + static_cast<std::size_t>(x) * kChannels,
                        PixelAt(source, sourceStride, sx, sy),
                        kChannels);
        }
    }
}

void ResampleBilinear(const std::uint8_t* source,
                      std::uint32_t sourceWidth,
                      std::uint32_t sourceHeight,
                      std::size_t sourceStride,
                      std::uint32_t targetWidth,
                      std::uint32_t targetHeight,
                      double scaleX,
                      double scaleY,
                      std::uint8_t* out) {
    /*
     * The two source columns and the horizontal weight depend only on x, so they
     * are computed once per column instead of once per pixel. On a 1920x1080
     * target that is a million redundant floor-and-clamp sequences removed, and
     * it costs one small table.
     */
    std::vector<std::pair<std::int64_t, std::int64_t>> columns(targetWidth);
    std::vector<float> columnWeights(targetWidth);
    for (std::uint32_t x = 0; x < targetWidth; ++x) {
        const double fx = SourceCoordinate(x, scaleX);
        const double baseX = std::floor(fx);
        columnWeights[x] = static_cast<float>(fx - baseX);
        columns[x].first = ClampCoordinate(static_cast<std::int64_t>(baseX), sourceWidth);
        columns[x].second = ClampCoordinate(static_cast<std::int64_t>(baseX) + 1, sourceWidth);
    }

    for (std::uint32_t y = 0; y < targetHeight; ++y) {
        const double fy = SourceCoordinate(y, scaleY);
        const double baseY = std::floor(fy);
        const float weightY = static_cast<float>(fy - baseY);
        const std::int64_t y0 = ClampCoordinate(static_cast<std::int64_t>(baseY), sourceHeight);
        const std::int64_t y1 = ClampCoordinate(static_cast<std::int64_t>(baseY) + 1, sourceHeight);

        std::uint8_t* row = out + static_cast<std::size_t>(y) * targetWidth * kChannels;
        for (std::uint32_t x = 0; x < targetWidth; ++x) {
            const std::int64_t x0 = columns[x].first;
            const std::int64_t x1 = columns[x].second;
            const float weightX = columnWeights[x];

            const Premultiplied p00 = LoadPremultiplied(PixelAt(source, sourceStride, x0, y0));
            const Premultiplied p10 = LoadPremultiplied(PixelAt(source, sourceStride, x1, y0));
            const Premultiplied p01 = LoadPremultiplied(PixelAt(source, sourceStride, x0, y1));
            const Premultiplied p11 = LoadPremultiplied(PixelAt(source, sourceStride, x1, y1));

            const float w00 = (1.0f - weightX) * (1.0f - weightY);
            const float w10 = weightX * (1.0f - weightY);
            const float w01 = (1.0f - weightX) * weightY;
            const float w11 = weightX * weightY;

            Premultiplied sum;
            sum.b = p00.b * w00 + p10.b * w10 + p01.b * w01 + p11.b * w11;
            sum.g = p00.g * w00 + p10.g * w10 + p01.g * w01 + p11.g * w11;
            sum.r = p00.r * w00 + p10.r * w10 + p01.r * w01 + p11.r * w11;
            sum.a = p00.a * w00 + p10.a * w10 + p01.a * w01 + p11.a * w11;
            StorePremultiplied(sum, row + static_cast<std::size_t>(x) * kChannels);
        }
    }
}

/**
 * Catmull-Rom, the cubic used here.
 *
 * It is the `a = -0.5` member of the Keys family, it interpolates - the curve
 * passes through the source samples, so an unscaled image is unchanged - and it
 * is the usual choice for image resampling. Mitchell-Netravali is smoother but
 * does not interpolate; B-spline blurs. Naming the exact kernel matters because
 * "bicubic" alone describes a family, not a result.
 */
inline float CatmullRom(float t) noexcept {
    const float x = std::fabs(t);
    if (x < 1.0f) {
        return 1.5f * x * x * x - 2.5f * x * x + 1.0f;
    }
    if (x < 2.0f) {
        return -0.5f * x * x * x + 2.5f * x * x - 4.0f * x + 2.0f;
    }
    return 0.0f;
}

void ResampleBicubic(const std::uint8_t* source,
                     std::uint32_t sourceWidth,
                     std::uint32_t sourceHeight,
                     std::size_t sourceStride,
                     std::uint32_t targetWidth,
                     std::uint32_t targetHeight,
                     double scaleX,
                     double scaleY,
                     std::uint8_t* out) {
    // Same reasoning as bilinear, and it matters more here: sixteen taps per
    // pixel each needed a weight that only ever depended on the column.
    std::vector<std::array<float, 4>> columnWeights(targetWidth);
    std::vector<std::array<std::int64_t, 4>> columnIndices(targetWidth);
    for (std::uint32_t x = 0; x < targetWidth; ++x) {
        const double fx = SourceCoordinate(x, scaleX);
        const double baseX = std::floor(fx);
        const float tx = static_cast<float>(fx - baseX);
        for (int i = 0; i < 4; ++i) {
            columnWeights[x][i] = CatmullRom(tx - static_cast<float>(i - 1));
            columnIndices[x][i] =
                ClampCoordinate(static_cast<std::int64_t>(baseX) + i - 1, sourceWidth);
        }
    }

    for (std::uint32_t y = 0; y < targetHeight; ++y) {
        const double fy = SourceCoordinate(y, scaleY);
        const double baseY = std::floor(fy);
        const float ty = static_cast<float>(fy - baseY);

        float weightsY[4];
        for (int i = 0; i < 4; ++i) {
            weightsY[i] = CatmullRom(ty - static_cast<float>(i - 1));
        }

        // The four source rows are the same for every pixel in this row.
        std::int64_t rows[4];
        for (int j = 0; j < 4; ++j) {
            rows[j] = ClampCoordinate(static_cast<std::int64_t>(baseY) + j - 1, sourceHeight);
        }

        std::uint8_t* row = out + static_cast<std::size_t>(y) * targetWidth * kChannels;
        for (std::uint32_t x = 0; x < targetWidth; ++x) {
            const std::array<float, 4>& weightsX = columnWeights[x];
            const std::array<std::int64_t, 4>& indicesX = columnIndices[x];

            Premultiplied sum;
            float total = 0.0f;
            for (int j = 0; j < 4; ++j) {
                const std::int64_t sy = rows[j];
                for (int i = 0; i < 4; ++i) {
                    const std::int64_t sx = indicesX[i];
                    const float weight = weightsX[i] * weightsY[j];
                    if (weight == 0.0f) {
                        continue;
                    }
                    const Premultiplied pixel =
                        LoadPremultiplied(PixelAt(source, sourceStride, sx, sy));
                    sum.b += pixel.b * weight;
                    sum.g += pixel.g * weight;
                    sum.r += pixel.r * weight;
                    sum.a += pixel.a * weight;
                    total += weight;
                }
            }
            // Catmull-Rom weights sum to 1 in exact arithmetic; normalising
            // removes the float drift that would otherwise darken or lighten a
            // flat region very slightly.
            if (total > 0.0f && std::fabs(total - 1.0f) > 1e-6f) {
                sum.b /= total;
                sum.g /= total;
                sum.r /= total;
                sum.a /= total;
            }
            StorePremultiplied(sum, row + static_cast<std::size_t>(x) * kChannels);
        }
    }
}

} // namespace

bool IsSupportedFilter(std::uint32_t filter) noexcept {
    switch (static_cast<ScaleFilter>(filter)) {
    case ScaleFilter::Nearest:
    case ScaleFilter::Bilinear:
    case ScaleFilter::Bicubic:
        return true;
    }
    return false;
}

const char* FilterName(ScaleFilter filter) noexcept {
    switch (filter) {
    case ScaleFilter::Nearest:
        return "nearest";
    case ScaleFilter::Bilinear:
        return "bilinear";
    case ScaleFilter::Bicubic:
        return "bicubic";
    }
    return "unknown";
}

const char* DescribeStatus(ResampleStatus status) noexcept {
    switch (status) {
    case ResampleStatus::Ok:
        return "ok";
    case ResampleStatus::NullSource:
        return "A source pixel buffer is required";
    case ResampleStatus::InvalidSourceSize:
        return "Source width and height must both be greater than zero";
    case ResampleStatus::InvalidTargetSize:
        return "Target width and height must both be greater than zero";
    case ResampleStatus::InvalidStride:
        return "Source stride must be at least width*4 bytes for BGRA32";
    case ResampleStatus::UnknownFilter:
        return "Unknown scaling filter";
    case ResampleStatus::TooLarge:
        return "Image exceeds the maximum supported size";
    case ResampleStatus::OutOfMemory:
        return "Could not allocate the resampled image";
    }
    return "Unknown resampling failure";
}

ResampleStatus Resample(const std::uint8_t* source,
                        std::uint32_t sourceWidth,
                        std::uint32_t sourceHeight,
                        std::int32_t sourceStride,
                        std::uint32_t targetWidth,
                        std::uint32_t targetHeight,
                        ScaleFilter filter,
                        std::vector<std::uint8_t>& out) {
    /*
     * Everything is validated before a byte is read or allocated, and every
     * product is formed in 64-bit and range-checked. A width and height that
     * overflowed a 32-bit size would otherwise produce a short allocation and a
     * long read, which is the classic way this kind of function becomes a
     * vulnerability.
     */
    if (!source) {
        return ResampleStatus::NullSource;
    }
    if (sourceWidth == 0 || sourceHeight == 0) {
        return ResampleStatus::InvalidSourceSize;
    }
    if (targetWidth == 0 || targetHeight == 0) {
        return ResampleStatus::InvalidTargetSize;
    }
    if (!IsSupportedFilter(static_cast<std::uint32_t>(filter))) {
        return ResampleStatus::UnknownFilter;
    }
    /*
     * Dimensions are fully validated before the buffer layout they imply. A
     * 65536x65536 source is nonsense whatever stride accompanies it, and
     * reporting "stride too small" for it would send a caller looking in the
     * wrong place.
     */
    const std::uint64_t sourcePixels =
        static_cast<std::uint64_t>(sourceWidth) * static_cast<std::uint64_t>(sourceHeight);
    const std::uint64_t targetPixels =
        static_cast<std::uint64_t>(targetWidth) * static_cast<std::uint64_t>(targetHeight);
    if (sourcePixels > kMaxPixels || targetPixels > kMaxPixels) {
        return ResampleStatus::TooLarge;
    }

    if (sourceStride <= 0) {
        return ResampleStatus::InvalidStride;
    }
    const std::uint64_t minimumStride = static_cast<std::uint64_t>(sourceWidth) * kChannels;
    if (static_cast<std::uint64_t>(sourceStride) < minimumStride) {
        return ResampleStatus::InvalidStride;
    }
    const std::uint64_t targetBytes = targetPixels * kChannels;
    if (targetBytes > static_cast<std::uint64_t>(SIZE_MAX)) {
        return ResampleStatus::TooLarge;
    }

    const std::size_t stride = static_cast<std::size_t>(sourceStride);
    try {
        out.resize(static_cast<std::size_t>(targetBytes));
    } catch (const std::bad_alloc&) {
        return ResampleStatus::OutOfMemory;
    } catch (...) {
        return ResampleStatus::OutOfMemory;
    }

    // Same size: copy the rows and run no filter. Cheaper, and it makes the
    // result bit-exact whichever filter was requested - which is what a caller
    // scaling "to the size it already is" should get.
    if (sourceWidth == targetWidth && sourceHeight == targetHeight) {
        const std::size_t rowBytes = static_cast<std::size_t>(sourceWidth) * kChannels;
        for (std::uint32_t y = 0; y < sourceHeight; ++y) {
            std::memcpy(out.data() + static_cast<std::size_t>(y) * rowBytes,
                        source + static_cast<std::size_t>(y) * stride,
                        rowBytes);
        }
        return ResampleStatus::Ok;
    }

    const double scaleX =
        static_cast<double>(sourceWidth) / static_cast<double>(targetWidth);
    const double scaleY =
        static_cast<double>(sourceHeight) / static_cast<double>(targetHeight);

    switch (filter) {
    case ScaleFilter::Nearest:
        ResampleNearest(source, sourceWidth, sourceHeight, stride, targetWidth, targetHeight,
                        scaleX, scaleY, out.data());
        break;
    case ScaleFilter::Bilinear:
        ResampleBilinear(source, sourceWidth, sourceHeight, stride, targetWidth, targetHeight,
                         scaleX, scaleY, out.data());
        break;
    case ScaleFilter::Bicubic:
        ResampleBicubic(source, sourceWidth, sourceHeight, stride, targetWidth, targetHeight,
                        scaleX, scaleY, out.data());
        break;
    }
    return ResampleStatus::Ok;
}

} // namespace bb::image
