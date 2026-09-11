/**
 * @file warp_contract.cpp
 * Correctness tests for the projective quad warp.
 *
 * The properties asserted here are the ones that separate a real homography from
 * an affine approximation that happens to look plausible:
 *
 *   - an axis-aligned rectangle warps to the identity, exactly
 *   - the four corners land on the four corners
 *   - a parallelogram stays affine (the projective terms vanish on their own)
 *   - a trapezoid foreshortens *non-linearly* - the source midpoint does not
 *     land at the quad's midpoint, which is precisely what an affine transform
 *     would get wrong
 *   - everything outside the quad is transparent
 *   - degenerate, duplicated and non-finite quads are refused, each by its own
 *     name, before anything is allocated
 *
 * Pure maths. No Office, no WIC, so it runs everywhere.
 */

#include "../src/image/warp.hpp"

#include <cmath>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>

namespace {

int g_failures = 0;

void Check(bool condition, const std::string& what) {
    if (condition) {
        std::printf("  ok   %s\n", what.c_str());
        return;
    }
    std::printf("  FAIL %s\n", what.c_str());
    ++g_failures;
}

using bb::image::PointF;
using bb::image::ScaleFilter;
using bb::image::WarpQuad;
using bb::image::WarpResult;
using bb::image::WarpStatus;

/// Pixel (x, y) carries its own coordinates, so a sample says where it came from.
std::vector<std::uint8_t> NumberedGrid(std::uint32_t size) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(size) * size * 4);
    for (std::uint32_t y = 0; y < size; ++y) {
        for (std::uint32_t x = 0; x < size; ++x) {
            auto* p = &pixels[(static_cast<std::size_t>(y) * size + x) * 4];
            p[0] = static_cast<std::uint8_t>(x);
            p[1] = static_cast<std::uint8_t>(y);
            p[2] = 0x40;
            p[3] = 255;
        }
    }
    return pixels;
}

struct Sample {
    int u, v, alpha;
};

Sample At(const bb::image::DecodedImage& image, std::uint32_t x, std::uint32_t y) {
    const auto* p = &image.pixels[(static_cast<std::size_t>(y) * image.width + x) * 4];
    return Sample{p[0], p[1], p[3]};
}

WarpResult Warp(const std::vector<std::uint8_t>& source,
                std::uint32_t size,
                const PointF quad[4],
                std::uint32_t width,
                std::uint32_t height,
                ScaleFilter filter = ScaleFilter::Nearest) {
    return WarpQuad(source.data(),
                    size,
                    size,
                    static_cast<std::int32_t>(size) * 4,
                    quad,
                    width,
                    height,
                    filter);
}

void TestRectangleIsIdentity() {
    constexpr std::uint32_t kSize = 64;
    const auto source = NumberedGrid(kSize);
    const PointF quad[4] = {{0, 0}, {64, 0}, {64, 64}, {0, 64}};

    const WarpResult result = Warp(source, kSize, quad, kSize, kSize);
    Check(result.ok(), "an axis-aligned rectangle warps");
    Check(result.image.width == kSize && result.image.height == kSize,
          "and keeps the requested raster size");
    Check(result.left == 0.0f && result.top == 0.0f && result.width == 64.0f &&
              result.height == 64.0f,
          "and reports the quad's bounding box unchanged");

    bool identity = true;
    for (std::uint32_t y = 0; y < kSize && identity; ++y) {
        for (std::uint32_t x = 0; x < kSize; ++x) {
            const Sample sample = At(result.image, x, y);
            if (sample.u != static_cast<int>(x) || sample.v != static_cast<int>(y) ||
                sample.alpha != 255) {
                identity = false;
                break;
            }
        }
    }
    Check(identity, "and every output pixel is exactly its own source pixel");
}

void TestCornersLandOnCorners() {
    constexpr std::uint32_t kSize = 64;
    const auto source = NumberedGrid(kSize);
    // A trapezoid: narrow at the top, wide at the bottom.
    const PointF quad[4] = {{40, 0}, {120, 0}, {160, 100}, {0, 100}};

    const WarpResult result = Warp(source, kSize, quad, 160, 100);
    Check(result.ok(), "a trapezoid warps");

    // Just inside each corner, because the corner pixel itself sits on the
    // boundary where a half-pixel either way decides in or out.
    const Sample topLeft = At(result.image, 42, 2);
    const Sample topRight = At(result.image, 117, 2);
    const Sample bottomLeft = At(result.image, 3, 97);
    const Sample bottomRight = At(result.image, 156, 97);

    Check(topLeft.alpha == 255 && topLeft.u < 8 && topLeft.v < 8,
          "the source top-left lands just inside the quad's first point");
    Check(topRight.alpha == 255 && topRight.u > 55 && topRight.v < 8,
          "the source top-right lands at the second point");
    Check(bottomRight.alpha == 255 && bottomRight.u > 55 && bottomRight.v > 55,
          "the source bottom-right lands at the third point");
    Check(bottomLeft.alpha == 255 && bottomLeft.u < 8 && bottomLeft.v > 55,
          "the source bottom-left lands at the fourth point");
}

void TestOutsideIsTransparent() {
    constexpr std::uint32_t kSize = 32;
    const auto source = NumberedGrid(kSize);
    const PointF quad[4] = {{50, 0}, {150, 0}, {200, 100}, {0, 100}};

    const WarpResult result = Warp(source, kSize, quad, 200, 100);
    Check(result.ok(), "a trapezoid inside a wider box warps");

    // The top corners of the bounding box are outside the trapezoid.
    Check(At(result.image, 2, 2).alpha == 0, "the top-left of the box is transparent");
    Check(At(result.image, 197, 2).alpha == 0, "the top-right of the box is transparent");
    Check(At(result.image, 100, 50).alpha == 255, "and the middle of the quad is painted");
}

void TestParallelogramStaysAffine() {
    constexpr std::uint32_t kSize = 64;
    const auto source = NumberedGrid(kSize);
    // Sheared, but opposite edges are still parallel, so the projective terms
    // must vanish and the mapping must be exactly linear.
    const PointF quad[4] = {{40, 0}, {140, 0}, {100, 100}, {0, 100}};

    const WarpResult result = Warp(source, kSize, quad, 140, 100);
    Check(result.ok(), "a parallelogram warps");

    // On an affine map the source centre lands at the quad's centroid.
    const float centroidX = (quad[0].x + quad[1].x + quad[2].x + quad[3].x) / 4.0f;
    const float centroidY = (quad[0].y + quad[1].y + quad[2].y + quad[3].y) / 4.0f;
    const Sample centre = At(result.image,
                             static_cast<std::uint32_t>(centroidX),
                             static_cast<std::uint32_t>(centroidY));
    Check(centre.alpha == 255 && std::abs(centre.u - 32) <= 2 && std::abs(centre.v - 32) <= 2,
          "the source centre lands at the centroid, so a parallelogram stayed affine");
}

void TestTrapezoidForeshortens() {
    constexpr std::uint32_t kSize = 64;
    const auto source = NumberedGrid(kSize);
    // Strongly tapered: the top edge is an eighth of the bottom edge, so
    // perspective and affine disagree by a lot.
    const PointF quad[4] = {{140, 0}, {180, 0}, {320, 200}, {0, 200}};

    const WarpResult result = Warp(source, kSize, quad, 320, 200);
    Check(result.ok(), "an extreme trapezoid warps");

    /*
     * The discriminating measurement.
     *
     * Down the quad's vertical centre line, find where the source's own
     * mid-row (v = 32 of 64) lands. An affine map puts it at exactly half the
     * height. A projective map puts it much nearer the narrow end, because the
     * far half of the texture is compressed into less space - which is what
     * perspective *is*.
     */
    int midRow = -1;
    for (std::uint32_t y = 0; y < result.image.height; ++y) {
        const Sample sample = At(result.image, result.image.width / 2, y);
        if (sample.alpha == 255 && sample.v >= 32) {
            midRow = static_cast<int>(y);
            break;
        }
    }
    Check(midRow > 0, "the source mid-row is somewhere on the centre line");
    Check(midRow > 0 && midRow < 85,
          "and it sits well above half height, so the mapping really is projective "
          "(affine would put it at 100)");

    // And the foreshortening has to be monotonic: v must never go backwards
    // walking down the centre line.
    int previous = -1;
    bool monotonic = true;
    for (std::uint32_t y = 0; y < result.image.height; ++y) {
        const Sample sample = At(result.image, result.image.width / 2, y);
        if (sample.alpha != 255) {
            continue;
        }
        if (sample.v < previous) {
            monotonic = false;
            break;
        }
        previous = sample.v;
    }
    Check(monotonic, "and v increases monotonically down the quad, with no folding");
}

void TestRefusals() {
    constexpr std::uint32_t kSize = 16;
    const auto source = NumberedGrid(kSize);

    const auto refused = [&](const PointF quad[4], WarpStatus expected, const char* what) {
        const WarpResult result = Warp(source, kSize, quad, 64, 64);
        Check(result.status == expected, what);
    };

    const PointF duplicate[4] = {{0, 0}, {0, 0}, {64, 64}, {0, 64}};
    refused(duplicate, WarpStatus::DuplicatePoints, "a quad with two identical corners is refused");

    const PointF collinear[4] = {{0, 0}, {10, 10}, {20, 20}, {30, 30}};
    refused(collinear, WarpStatus::DegenerateQuad, "four collinear points are refused");

    // A very thin rectangle is not degenerate - it has a perfectly good inverse
    // and a caller asking for a 64-by-nothing sliver should get one. Refusing it
    // would be this file inventing a minimum size nobody asked for. What must be
    // refused is a quad with no *interior*, which the collinear case above is.
    const PointF thin[4] = {{0, 0}, {64, 0}, {64, 0.0001f}, {0, 0.0001f}};
    const WarpResult sliver = Warp(source, kSize, thin, 0, 0);
    Check(sliver.ok() && sliver.image.height >= 1,
          "a thin but non-degenerate quad is accepted, not refused for being small");

    const PointF bowtie[4] = {{0, 0}, {64, 64}, {64, 0}, {0, 64}};
    const WarpResult crossed = Warp(source, kSize, bowtie, 64, 64);
    Check(crossed.status == WarpStatus::DegenerateQuad,
          "a self-crossing bow-tie encloses no net area and is refused");

    const PointF nan[4] = {
        {0, 0}, {std::nanf(""), 0}, {64, 64}, {0, 64}};
    refused(nan, WarpStatus::NonFinitePoint, "a not-a-number coordinate is refused");

    const PointF infinite[4] = {
        {0, 0}, {std::numeric_limits<float>::infinity(), 0}, {64, 64}, {0, 64}};
    refused(infinite, WarpStatus::NonFinitePoint, "an infinite coordinate is refused");

    const PointF fine[4] = {{0, 0}, {64, 0}, {64, 64}, {0, 64}};
    const WarpResult huge = Warp(source, kSize, fine, 40000, 40000);
    Check(huge.status == WarpStatus::TooLarge, "an output raster over 64 megapixels is refused");

    const WarpResult bicubic = Warp(source, kSize, fine, 32, 32, ScaleFilter::Bicubic);
    Check(bicubic.status == WarpStatus::UnknownFilter,
          "bicubic is refused by name rather than silently downgraded");

    const WarpResult noSource =
        WarpQuad(nullptr, 16, 16, 64, fine, 32, 32, ScaleFilter::Nearest);
    Check(noSource.status == WarpStatus::NullSource, "a null source is refused");
}

void TestPointOrderIsNotReordered() {
    constexpr std::uint32_t kSize = 32;
    const auto source = NumberedGrid(kSize);

    // The same rectangle, with the top two corners swapped. That is a caller
    // asking for a mirrored mapping, and they must get one rather than have the
    // points quietly sorted back into place.
    const PointF mirrored[4] = {{64, 0}, {0, 0}, {0, 64}, {64, 64}};
    const WarpResult result = Warp(source, kSize, mirrored, 64, 64);
    Check(result.ok(), "a deliberately mirrored quad warps");

    const Sample left = At(result.image, 4, 4);
    const Sample right = At(result.image, 59, 4);
    Check(left.alpha == 255 && right.alpha == 255, "both top corners are painted");
    Check(left.u > right.u,
          "and the mapping is mirrored, so the points were not reordered behind the caller");
}

} // namespace

int main() {
    std::printf("warp contract\n");
    TestRectangleIsIdentity();
    TestCornersLandOnCorners();
    TestOutsideIsTransparent();
    TestParallelogramStaysAffine();
    TestTrapezoidForeshortens();
    TestRefusals();
    TestPointOrderIsNotReordered();

    if (g_failures != 0) {
        std::printf("%d check(s) failed\n", g_failures);
        return 1;
    }
    std::printf("all checks passed\n");
    return 0;
}
