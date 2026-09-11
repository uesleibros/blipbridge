/**
 * @file pipeline_contract.cpp
 * Correctness tests for crop, orient and resize, and for the order they run in.
 *
 * Pure image maths plus WIC, no Office, so this runs on both architectures.
 *
 * Every fixture is a numbered grid: pixel (x, y) carries its own coordinates in
 * its blue and green channels. That makes every assertion here a statement about
 * *where a pixel came from* rather than about what colour it happens to be, so a
 * transpose, a mirror or an off-by-one shows up as a specific wrong coordinate
 * instead of as a vague difference.
 */

#include "../src/image/pipeline.hpp"

#include <algorithm>
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

using bb::image::DecodedImage;
using bb::image::PrepareRequest;
using bb::image::PrepareResult;
using bb::image::PrepareStatus;
using bb::image::ScaleFilter;
using bb::image::Transform;

/// Pixel (x, y) is B=x, G=y, R=a constant, A=255. Its own coordinates, in it.
std::vector<std::uint8_t> NumberedGrid(std::uint32_t width, std::uint32_t height) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(width) * height * 4);
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            auto* p = &pixels[(static_cast<std::size_t>(y) * width + x) * 4];
            p[0] = static_cast<std::uint8_t>(x);
            p[1] = static_cast<std::uint8_t>(y);
            p[2] = 0x40;
            p[3] = 255;
        }
    }
    return pixels;
}

/// Where the pixel now at (x, y) came from in the source grid.
struct Origin {
    int x;
    int y;
};

Origin OriginAt(const DecodedImage& image, std::uint32_t x, std::uint32_t y) {
    if (x >= image.width || y >= image.height) {
        return Origin{-1, -1};
    }
    const auto* p = &image.pixels[(static_cast<std::size_t>(y) * image.width + x) * 4];
    return Origin{p[0], p[1]};
}

bool Is(const DecodedImage& image, std::uint32_t x, std::uint32_t y, int sourceX, int sourceY) {
    const Origin origin = OriginAt(image, x, y);
    return origin.x == sourceX && origin.y == sourceY;
}

PrepareResult Run(const std::vector<std::uint8_t>& source,
                  std::uint32_t width,
                  std::uint32_t height,
                  const PrepareRequest& request) {
    return bb::image::PreparePixels(
        source.data(), width, height, static_cast<std::int32_t>(width) * 4, request);
}

void TestCrop() {
    constexpr std::uint32_t kSize = 64;
    const auto source = NumberedGrid(kSize, kSize);

    {
        // A 16x16 tile out of an atlas, the motivating case.
        PrepareRequest request;
        request.cropX = 16;
        request.cropY = 32;
        request.cropWidth = 16;
        request.cropHeight = 16;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok(), "a 16x16 tile crops");
        Check(result.image.width == 16 && result.image.height == 16, "and is 16x16");
        Check(Is(result.image, 0, 0, 16, 32), "its top-left is the source pixel at (16,32)");
        Check(Is(result.image, 15, 15, 31, 47), "its bottom-right is the source pixel at (31,47)");
    }
    {
        PrepareRequest request;
        request.cropX = 63;
        request.cropY = 63;
        request.cropWidth = 1;
        request.cropHeight = 1;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok() && result.image.width == 1 && result.image.height == 1,
              "a 1x1 region at the far corner crops");
        Check(Is(result.image, 0, 0, 63, 63), "and is exactly that pixel");
    }
    {
        PrepareRequest request;
        request.cropWidth = kSize;
        request.cropHeight = kSize;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok() && result.image.pixels == source,
              "a full-image region is the image, byte for byte");
    }
    {
        const PrepareResult result = Run(source, kSize, kSize, PrepareRequest{});
        Check(result.ok() && result.image.pixels == source,
              "a default request changes nothing at all");
    }
}

void TestCropRefusals() {
    constexpr std::uint32_t kSize = 32;
    const auto source = NumberedGrid(kSize, kSize);

    const auto refused = [&](PrepareRequest request, const char* what) {
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.status == PrepareStatus::InvalidCrop, what);
    };

    PrepareRequest past;
    past.cropX = 16;
    past.cropWidth = 17;
    past.cropHeight = 4;
    refused(past, "a region running off the right edge is refused");

    PrepareRequest below;
    below.cropY = 30;
    below.cropWidth = 4;
    below.cropHeight = 4;
    refused(below, "a region running off the bottom edge is refused");

    PrepareRequest outside;
    outside.cropX = kSize;
    outside.cropWidth = 1;
    outside.cropHeight = 1;
    refused(outside, "a region starting outside the image is refused");

    // The overflow case: an origin near the top of the range plus a width that
    // would wrap it back inside if either were computed in 32 bits.
    PrepareRequest overflow;
    overflow.cropX = 0xFFFFFFFFu;
    overflow.cropY = 0xFFFFFFFFu;
    overflow.cropWidth = 8;
    overflow.cropHeight = 8;
    refused(overflow, "a region whose origin plus size would overflow 32 bits is refused");

    PrepareRequest originOnly;
    originOnly.cropX = 4;
    originOnly.cropY = 4;
    refused(originOnly, "an origin with no size is refused rather than silently ignored");
}

void TestTransforms() {
    // 4 wide, 2 tall, so a transpose is visible in the dimensions alone.
    constexpr std::uint32_t kWidth = 4;
    constexpr std::uint32_t kHeight = 2;
    const auto source = NumberedGrid(kWidth, kHeight);

    const auto run = [&](Transform transform) {
        PrepareRequest request;
        request.transform = transform;
        return Run(source, kWidth, kHeight, request);
    };

    {
        const PrepareResult result = run(Transform::FlipHorizontal);
        Check(result.ok() && result.image.width == kWidth && result.image.height == kHeight,
              "a horizontal flip keeps the dimensions");
        Check(Is(result.image, 0, 0, 3, 0) && Is(result.image, 3, 0, 0, 0),
              "and mirrors the top row left to right");
        Check(Is(result.image, 0, 1, 3, 1), "and does the same to every row");
    }
    {
        const PrepareResult result = run(Transform::FlipVertical);
        Check(result.ok() && result.image.width == kWidth && result.image.height == kHeight,
              "a vertical flip keeps the dimensions");
        Check(Is(result.image, 0, 0, 0, 1) && Is(result.image, 0, 1, 0, 0),
              "and swaps top for bottom without mirroring columns");
    }
    {
        const PrepareResult result = run(Transform::Rotate90);
        Check(result.ok() && result.image.width == kHeight && result.image.height == kWidth,
              "rotating 90 transposes the dimensions to 2x4");
        // Clockwise: the source top-left ends up at the top-right.
        Check(Is(result.image, 1, 0, 0, 0), "the source top-left lands top-right, so it is clockwise");
        Check(Is(result.image, 1, 3, 3, 0), "and the source top-right lands bottom-right");
        Check(Is(result.image, 0, 0, 0, 1), "with the second row becoming the left column");
    }
    {
        const PrepareResult result = run(Transform::Rotate180);
        Check(result.ok() && result.image.width == kWidth && result.image.height == kHeight,
              "rotating 180 keeps the dimensions");
        Check(Is(result.image, 0, 0, 3, 1) && Is(result.image, 3, 1, 0, 0),
              "and maps each corner to the opposite one");
    }
    {
        const PrepareResult result = run(Transform::Rotate270);
        Check(result.ok() && result.image.width == kHeight && result.image.height == kWidth,
              "rotating 270 transposes the dimensions to 2x4");
        Check(Is(result.image, 0, 3, 0, 0), "the source top-left lands bottom-left");
        Check(Is(result.image, 0, 0, 3, 0), "and the source top-right lands top-left");
    }
    {
        // Two opposite rotations must return the original exactly.
        PrepareRequest first;
        first.transform = Transform::Rotate90;
        const PrepareResult once = Run(source, kWidth, kHeight, first);
        PrepareRequest second;
        second.transform = Transform::Rotate270;
        const PrepareResult back = bb::image::PreparePixels(
            once.image.pixels.data(), once.image.width, once.image.height,
            once.image.stride(), second);
        Check(back.ok() && back.image.pixels == source,
              "rotating 90 then 270 returns the original, byte for byte");
    }
    {
        PrepareRequest unknown;
        unknown.transform = static_cast<Transform>(99);
        const PrepareResult result = Run(source, kWidth, kHeight, unknown);
        Check(result.status == PrepareStatus::UnknownTransform,
              "a transform this build does not implement is refused by name");
    }
}

void TestStageOrder() {
    constexpr std::uint32_t kSize = 16;
    const auto source = NumberedGrid(kSize, kSize);

    // Crop a 4x2 region, rotate it 90, then ask for 8x16. If resize ran before
    // the rotation, the target would describe an image nobody asked for and the
    // result would come out 16x8.
    PrepareRequest request;
    request.cropX = 2;
    request.cropY = 4;
    request.cropWidth = 4;
    request.cropHeight = 2;
    request.transform = Transform::Rotate90;
    request.targetWidth = 8;
    request.targetHeight = 16;
    request.filter = ScaleFilter::Nearest;

    const PrepareResult result = Run(source, kSize, kSize, request);
    Check(result.ok(), "crop, rotate and resize together succeed");
    Check(result.image.width == 8 && result.image.height == 16,
          "the target size describes the final image, so the rotation ran before the resize");

    // Nearest, so the corner pixel is still exactly a source pixel: the rotated
    // crop's top-right is the crop's top-left, which is source (2,4).
    Check(Is(result.image, result.image.width - 1, 0, 2, 4),
          "and nearest kept the corner pixel exactly, from the cropped and rotated position");
}

void TestScaling() {
    constexpr std::uint32_t kSize = 8;
    const auto source = NumberedGrid(kSize, kSize);

    {
        PrepareRequest request;
        request.targetWidth = kSize;
        request.targetHeight = kSize;
        request.filter = ScaleFilter::Bicubic;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok() && result.image.pixels == source,
              "asking for the size it already is returns it bit-exact, whatever the filter");
    }
    {
        PrepareRequest request;
        request.targetWidth = kSize * 2;
        request.targetHeight = kSize * 2;
        request.filter = ScaleFilter::Nearest;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok() && result.image.width == 16 && result.image.height == 16,
              "a nearest double produces the target size");
        Check(Is(result.image, 0, 0, 0, 0) && Is(result.image, 1, 1, 0, 0) &&
                  Is(result.image, 2, 2, 1, 1),
              "and every source pixel became an exact 2x2 block");
    }
    {
        // Crop then scale, the atlas case end to end.
        PrepareRequest request;
        request.cropX = 4;
        request.cropY = 0;
        request.cropWidth = 4;
        request.cropHeight = 4;
        request.targetWidth = 8;
        request.targetHeight = 8;
        request.filter = ScaleFilter::Nearest;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.ok() && result.image.width == 8, "crop then nearest-scale succeeds");
        Check(Is(result.image, 0, 0, 4, 0) && Is(result.image, 7, 7, 7, 3),
              "and the scaled tile is the cropped region, not the whole image");
    }
    {
        PrepareRequest request;
        request.targetWidth = 0xFFFFFFFFu;
        request.targetHeight = 0xFFFFFFFFu;
        const PrepareResult result = Run(source, kSize, kSize, request);
        Check(result.status == PrepareStatus::ResampleFailed,
              "an absurd target size is refused by the resampler rather than attempted");
    }
}

void TestAlphaSurvives() {
    // A checkerboard of opaque and fully transparent, which is what a cut-out
    // sprite looks like and what a premultiply mistake would smear.
    constexpr std::uint32_t kSize = 8;
    std::vector<std::uint8_t> source(static_cast<std::size_t>(kSize) * kSize * 4);
    for (std::uint32_t y = 0; y < kSize; ++y) {
        for (std::uint32_t x = 0; x < kSize; ++x) {
            auto* p = &source[(static_cast<std::size_t>(y) * kSize + x) * 4];
            const bool solid = ((x + y) % 2) == 0;
            p[0] = 0;
            p[1] = 0;
            p[2] = solid ? 255 : 0;
            p[3] = solid ? 255 : 0;
        }
    }

    PrepareRequest request;
    request.transform = Transform::Rotate180;
    const PrepareResult result = Run(source, kSize, kSize, request);
    Check(result.ok(), "a transparent checkerboard rotates");

    bool preserved = true;
    for (std::uint32_t y = 0; y < kSize && preserved; ++y) {
        for (std::uint32_t x = 0; x < kSize; ++x) {
            const auto* p = &result.image.pixels[(static_cast<std::size_t>(y) * kSize + x) * 4];
            const bool solid = ((x + y) % 2) == 0;
            if (p[3] != (solid ? 255 : 0)) {
                preserved = false;
                break;
            }
        }
    }
    Check(preserved, "and every pixel kept its alpha exactly - a transform never blends");
}

} // namespace

int main() {
    std::printf("pipeline contract\n");
    TestCrop();
    TestCropRefusals();
    TestTransforms();
    TestStageOrder();
    TestScaling();
    TestAlphaSurvives();

    if (g_failures != 0) {
        std::printf("%d check(s) failed\n", g_failures);
        return 1;
    }
    std::printf("all checks passed\n");
    return 0;
}
