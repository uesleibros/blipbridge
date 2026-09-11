/**
 * @file warp.cpp
 * The projective sampler. See warp.hpp for why pre-warping is viable at all.
 */

#include "warp.hpp"

#include <cmath>
#include <cstring>
#include <new>

namespace bb::image {
namespace {

/**
 * A 3x3 projective transform, row-major.
 *
 * Only ever built by SolveUnitSquareTo and inverted by Invert, so there is no
 * general matrix type here - a general one would invite general uses, and this
 * needs exactly two operations.
 */
struct Matrix3 {
    double m[9]{};
};

/**
 * The homography taking the unit square's corners to @p quad.
 *
 * The standard construction: map the unit square to the quad directly rather
 * than composing two transforms, which keeps the arithmetic short enough to read
 * and avoids an intermediate inverse.
 *
 * With corners q0..q3 at (0,0), (1,0), (1,1), (0,1), the map is
 *
 *     [ a b c ]
 *     [ d e f ]
 *     [ g h 1 ]
 *
 * and the two projective terms g and h fall out of how far the quad is from a
 * parallelogram. When it *is* a parallelogram both are zero and this degenerates
 * to the affine case on its own, with no special path.
 */
bool SolveUnitSquareTo(const PointF quad[4], Matrix3& out) noexcept {
    const double x0 = quad[0].x, y0 = quad[0].y;
    const double x1 = quad[1].x, y1 = quad[1].y;
    const double x2 = quad[2].x, y2 = quad[2].y;
    const double x3 = quad[3].x, y3 = quad[3].y;

    const double dx1 = x1 - x2;
    const double dx2 = x3 - x2;
    const double dy1 = y1 - y2;
    const double dy2 = y3 - y2;
    const double sx = x0 - x1 + x2 - x3;
    const double sy = y0 - y1 + y2 - y3;

    double g = 0.0;
    double h = 0.0;
    const double denominator = dx1 * dy2 - dx2 * dy1;
    if (std::fabs(denominator) < 1e-12) {
        // The two edge vectors are parallel or one is zero: no projective
        // solution exists, and pretending otherwise would divide by noise.
        return false;
    }
    if (std::fabs(sx) > 1e-12 || std::fabs(sy) > 1e-12) {
        g = (sx * dy2 - dx2 * sy) / denominator;
        h = (dx1 * sy - sx * dy1) / denominator;
    }

    out.m[0] = x1 - x0 + g * x1;
    out.m[1] = x3 - x0 + h * x3;
    out.m[2] = x0;
    out.m[3] = y1 - y0 + g * y1;
    out.m[4] = y3 - y0 + h * y3;
    out.m[5] = y0;
    out.m[6] = g;
    out.m[7] = h;
    out.m[8] = 1.0;
    return true;
}

/// Inverts @p in, or reports that it has no inverse.
bool Invert(const Matrix3& in, Matrix3& out) noexcept {
    const double* m = in.m;
    const double c00 = m[4] * m[8] - m[5] * m[7];
    const double c01 = m[5] * m[6] - m[3] * m[8];
    const double c02 = m[3] * m[7] - m[4] * m[6];
    const double determinant = m[0] * c00 + m[1] * c01 + m[2] * c02;
    if (!std::isfinite(determinant) || std::fabs(determinant) < 1e-12) {
        return false;
    }
    const double inverse = 1.0 / determinant;
    out.m[0] = c00 * inverse;
    out.m[1] = (m[2] * m[7] - m[1] * m[8]) * inverse;
    out.m[2] = (m[1] * m[5] - m[2] * m[4]) * inverse;
    out.m[3] = c01 * inverse;
    out.m[4] = (m[0] * m[8] - m[2] * m[6]) * inverse;
    out.m[5] = (m[2] * m[3] - m[0] * m[5]) * inverse;
    out.m[6] = c02 * inverse;
    out.m[7] = (m[1] * m[6] - m[0] * m[7]) * inverse;
    out.m[8] = (m[0] * m[4] - m[1] * m[3]) * inverse;
    return true;
}

struct Bgra {
    double b, g, r, a;
};

/// One source texel, clamped to the edge like every other filter here.
Bgra Texel(const std::uint8_t* source,
           std::int32_t stride,
           std::uint32_t width,
           std::uint32_t height,
           long x,
           long y) noexcept {
    if (x < 0) x = 0;
    if (y < 0) y = 0;
    if (x >= static_cast<long>(width)) x = static_cast<long>(width) - 1;
    if (y >= static_cast<long>(height)) y = static_cast<long>(height) - 1;
    const std::uint8_t* p = source + static_cast<std::size_t>(y) * stride + static_cast<std::size_t>(x) * 4;
    return Bgra{static_cast<double>(p[0]),
                static_cast<double>(p[1]),
                static_cast<double>(p[2]),
                static_cast<double>(p[3])};
}

/// Bilinear in premultiplied alpha, so transparent texels cannot darken edges.
Bgra SampleBilinear(const std::uint8_t* source,
                    std::int32_t stride,
                    std::uint32_t width,
                    std::uint32_t height,
                    double u,
                    double v) noexcept {
    const double x = u * width - 0.5;
    const double y = v * height - 0.5;
    const auto x0 = static_cast<long>(std::floor(x));
    const auto y0 = static_cast<long>(std::floor(y));
    const double fx = x - x0;
    const double fy = y - y0;

    const Bgra corners[4] = {Texel(source, stride, width, height, x0, y0),
                             Texel(source, stride, width, height, x0 + 1, y0),
                             Texel(source, stride, width, height, x0, y0 + 1),
                             Texel(source, stride, width, height, x0 + 1, y0 + 1)};
    const double weights[4] = {
        (1.0 - fx) * (1.0 - fy), fx * (1.0 - fy), (1.0 - fx) * fy, fx * fy};

    Bgra sum{0, 0, 0, 0};
    for (int index = 0; index < 4; ++index) {
        const double alpha = corners[index].a / 255.0;
        sum.b += corners[index].b * alpha * weights[index];
        sum.g += corners[index].g * alpha * weights[index];
        sum.r += corners[index].r * alpha * weights[index];
        sum.a += corners[index].a * weights[index];
    }
    if (sum.a > 0.0) {
        const double unpremultiply = 255.0 / sum.a;
        sum.b *= unpremultiply;
        sum.g *= unpremultiply;
        sum.r *= unpremultiply;
    }
    return sum;
}

std::uint8_t Clamp255(double value) noexcept {
    if (!(value > 0.0)) {
        return 0;
    }
    if (value >= 255.0) {
        return 255;
    }
    return static_cast<std::uint8_t>(value + 0.5);
}

} // namespace

const char* DescribeWarpStatus(WarpStatus status) noexcept {
    switch (status) {
    case WarpStatus::Ok:
        return "The image was warped";
    case WarpStatus::NullSource:
        return "No source image was given";
    case WarpStatus::InvalidSourceSize:
        return "The source image has no pixels, or an impossible stride";
    case WarpStatus::NonFinitePoint:
        return "A quad corner is infinite or not a number";
    case WarpStatus::DuplicatePoints:
        return "Two or more quad corners are the same point";
    case WarpStatus::DegenerateQuad:
        return "The quad encloses no area";
    case WarpStatus::TooLarge:
        return "The quad's bounding box is larger than 64 megapixels";
    case WarpStatus::NotInvertible:
        return "The quad has no projective mapping - its corners are collinear";
    case WarpStatus::UnknownFilter:
        return "Warping supports nearest and bilinear only";
    case WarpStatus::OutOfMemory:
        return "Not enough memory to hold the warped image";
    }
    return "Unknown warp status";
}

WarpResult WarpQuad(const std::uint8_t* source,
                    std::uint32_t sourceWidth,
                    std::uint32_t sourceHeight,
                    std::int32_t sourceStride,
                    const PointF quad[4],
                    std::uint32_t outputWidth,
                    std::uint32_t outputHeight,
                    ScaleFilter filter) noexcept {
    WarpResult result;

    if (!source || !quad) {
        result.status = WarpStatus::NullSource;
        return result;
    }
    if (sourceWidth == 0 || sourceHeight == 0 ||
        sourceStride < static_cast<std::int32_t>(sourceWidth) * 4) {
        result.status = WarpStatus::InvalidSourceSize;
        return result;
    }
    if (filter != ScaleFilter::Nearest && filter != ScaleFilter::Bilinear) {
        result.status = WarpStatus::UnknownFilter;
        return result;
    }

    // Every refusal below happens before a byte is allocated, and each one says
    // something different: a caller who passed the same point twice has a
    // different bug from one who passed a NaN.
    for (int index = 0; index < 4; ++index) {
        if (!std::isfinite(quad[index].x) || !std::isfinite(quad[index].y)) {
            result.status = WarpStatus::NonFinitePoint;
            return result;
        }
    }
    for (int a = 0; a < 4; ++a) {
        for (int b = a + 1; b < 4; ++b) {
            if (quad[a].x == quad[b].x && quad[a].y == quad[b].y) {
                result.status = WarpStatus::DuplicatePoints;
                return result;
            }
        }
    }

    float left = quad[0].x;
    float top = quad[0].y;
    float right = quad[0].x;
    float bottom = quad[0].y;
    for (int index = 1; index < 4; ++index) {
        left = quad[index].x < left ? quad[index].x : left;
        top = quad[index].y < top ? quad[index].y : top;
        right = quad[index].x > right ? quad[index].x : right;
        bottom = quad[index].y > bottom ? quad[index].y : bottom;
    }
    const double boxWidth = static_cast<double>(right) - left;
    const double boxHeight = static_cast<double>(bottom) - top;
    if (!(boxWidth > 0.0) || !(boxHeight > 0.0)) {
        result.status = WarpStatus::DegenerateQuad;
        return result;
    }

    // The shoelace area, which catches the quads a bounding box cannot: four
    // distinct but collinear points have a box and no interior.
    double area = 0.0;
    for (int index = 0; index < 4; ++index) {
        const PointF& current = quad[index];
        const PointF& next = quad[(index + 1) % 4];
        area += static_cast<double>(current.x) * next.y - static_cast<double>(next.x) * current.y;
    }
    area = std::fabs(area) * 0.5;
    if (area < boxWidth * boxHeight * 1e-6) {
        result.status = WarpStatus::DegenerateQuad;
        return result;
    }

    std::uint32_t width = outputWidth;
    std::uint32_t height = outputHeight;
    if (width == 0 || height == 0) {
        // One output pixel per unit of the caller's own space. Rounded up, so a
        // fractional box never loses its last column.
        const double w = std::ceil(boxWidth);
        const double h = std::ceil(boxHeight);
        if (w < 1.0 || h < 1.0 || w > 65535.0 || h > 65535.0) {
            result.status = WarpStatus::TooLarge;
            return result;
        }
        width = static_cast<std::uint32_t>(w);
        height = static_cast<std::uint32_t>(h);
    }
    const std::uint64_t pixels = static_cast<std::uint64_t>(width) * height;
    if (pixels == 0 || pixels > kMaxPixels) {
        result.status = WarpStatus::TooLarge;
        return result;
    }

    /*
     * The quad in output-raster coordinates, then the transform.
     *
     * Solving unit-square -> quad and inverting gives the map from an output
     * pixel back to a source (u, v), which is the direction a sampler needs: one
     * inverse mapping per output pixel, no subdivision of the source, and no
     * holes where a forward map would leave them.
     */
    PointF raster[4];
    for (int index = 0; index < 4; ++index) {
        raster[index].x =
            static_cast<float>((quad[index].x - left) / boxWidth * width);
        raster[index].y =
            static_cast<float>((quad[index].y - top) / boxHeight * height);
    }

    Matrix3 forward;
    Matrix3 inverse;
    if (!SolveUnitSquareTo(raster, forward) || !Invert(forward, inverse)) {
        result.status = WarpStatus::NotInvertible;
        return result;
    }

    DecodedImage image;
    image.width = width;
    image.height = height;
    try {
        image.pixels.assign(static_cast<std::size_t>(pixels) * 4, 0);
    } catch (const std::bad_alloc&) {
        result.status = WarpStatus::OutOfMemory;
        return result;
    }

    const double* inv = inverse.m;
    for (std::uint32_t y = 0; y < height; ++y) {
        const double py = y + 0.5;
        auto* row = &image.pixels[static_cast<std::size_t>(y) * width * 4];
        for (std::uint32_t x = 0; x < width; ++x) {
            const double px = x + 0.5;
            const double w = inv[6] * px + inv[7] * py + inv[8];
            if (std::fabs(w) < 1e-12) {
                continue; // on the horizon of the projection; nothing maps here
            }
            const double u = (inv[0] * px + inv[1] * py + inv[2]) / w;
            const double v = (inv[3] * px + inv[4] * py + inv[5]) / w;

            // Outside the unit square is outside the quad. Left transparent, so
            // the result is right whether or not the Shape's path also clips.
            if (u < 0.0 || u >= 1.0 || v < 0.0 || v >= 1.0) {
                continue;
            }

            std::uint8_t* out = row + static_cast<std::size_t>(x) * 4;
            if (filter == ScaleFilter::Nearest) {
                auto sx = static_cast<long>(u * sourceWidth);
                auto sy = static_cast<long>(v * sourceHeight);
                const Bgra texel = Texel(source, sourceStride, sourceWidth, sourceHeight, sx, sy);
                out[0] = static_cast<std::uint8_t>(texel.b);
                out[1] = static_cast<std::uint8_t>(texel.g);
                out[2] = static_cast<std::uint8_t>(texel.r);
                out[3] = static_cast<std::uint8_t>(texel.a);
            } else {
                const Bgra texel =
                    SampleBilinear(source, sourceStride, sourceWidth, sourceHeight, u, v);
                out[0] = Clamp255(texel.b);
                out[1] = Clamp255(texel.g);
                out[2] = Clamp255(texel.r);
                out[3] = Clamp255(texel.a);
            }
        }
    }

    result.image = std::move(image);
    result.left = left;
    result.top = top;
    result.width = static_cast<float>(boxWidth);
    result.height = static_cast<float>(boxHeight);
    return result;
}

} // namespace bb::image
