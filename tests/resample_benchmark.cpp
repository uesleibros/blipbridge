/**
 * @file resample_benchmark.cpp
 * Times each resampling filter, so a claim about cost is a measurement.
 *
 * Scaling only - no Office, no cached-image creation, no apply. Those are timed
 * separately by the in-PowerPoint harnesses, and mixing them here would hide
 * which half a number belonged to.
 *
 * Built as a normal executable rather than a CTest case: a timing run is not a
 * pass/fail assertion and should not be able to fail CI because a hosted runner
 * was busy.
 */

#include "../src/image/resample.hpp"

#include <chrono>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

namespace {

using Clock = std::chrono::steady_clock;

/// A source image with a bit of structure, so no filter can shortcut on a flat
/// field and look faster than it is.
std::vector<std::uint8_t> MakeSource(std::uint32_t width, std::uint32_t height) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(width) * height * 4);
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            std::uint8_t* pixel = pixels.data() + (static_cast<std::size_t>(y) * width + x) * 4;
            pixel[0] = static_cast<std::uint8_t>(x * 7 + y);
            pixel[1] = static_cast<std::uint8_t>(y * 5 + x);
            pixel[2] = static_cast<std::uint8_t>((x ^ y) * 3);
            pixel[3] = static_cast<std::uint8_t>(x % 2 == 0 ? 255 : 128);
        }
    }
    return pixels;
}

double TimeOne(const std::vector<std::uint8_t>& source,
               std::uint32_t width,
               std::uint32_t height,
               std::uint32_t targetWidth,
               std::uint32_t targetHeight,
               bb::image::ScaleFilter filter,
               int iterations) {
    std::vector<std::uint8_t> out;
    // One warm-up: the first call pays for the allocation every later one reuses.
    bb::image::Resample(source.data(), width, height, static_cast<std::int32_t>(width) * 4,
                        targetWidth, targetHeight, filter, out);

    const Clock::time_point start = Clock::now();
    for (int i = 0; i < iterations; ++i) {
        bb::image::Resample(source.data(), width, height, static_cast<std::int32_t>(width) * 4,
                            targetWidth, targetHeight, filter, out);
    }
    const Clock::time_point end = Clock::now();
    const double totalMs = std::chrono::duration<double, std::milli>(end - start).count();
    return totalMs / iterations;
}

struct Case {
    const char* name;
    std::uint32_t width;
    std::uint32_t height;
    std::uint32_t targetWidth;
    std::uint32_t targetHeight;
    int iterations;
};

} // namespace

int main() {
    const Case cases[] = {
        {"small upscale   32x32 -> 128x128", 32, 32, 128, 128, 200},
        {"small downscale 128x128 -> 32x32", 128, 128, 32, 32, 200},
        {"medium upscale  256x256 -> 512x512", 256, 256, 512, 512, 50},
        {"medium downscale 512x512 -> 256x256", 512, 512, 256, 256, 50},
        {"large upscale   512x512 -> 1920x1080", 512, 512, 1920, 1080, 20},
        {"large downscale 1920x1080 -> 512x512", 1920, 1080, 512, 512, 20},
        {"non-integer     100x75 -> 333x251", 100, 75, 333, 251, 50},
        {"identity        512x512 -> 512x512", 512, 512, 512, 512, 50},
    };

    std::printf("%-38s %10s %10s %10s\n", "case", "nearest", "bilinear", "bicubic");
    std::printf("%-38s %10s %10s %10s\n", "----", "-------", "--------", "-------");

    for (const Case& one : cases) {
        const std::vector<std::uint8_t> source = MakeSource(one.width, one.height);
        const double nearest = TimeOne(source, one.width, one.height, one.targetWidth,
                                       one.targetHeight, bb::image::ScaleFilter::Nearest,
                                       one.iterations);
        const double bilinear = TimeOne(source, one.width, one.height, one.targetWidth,
                                        one.targetHeight, bb::image::ScaleFilter::Bilinear,
                                        one.iterations);
        const double bicubic = TimeOne(source, one.width, one.height, one.targetWidth,
                                       one.targetHeight, bb::image::ScaleFilter::Bicubic,
                                       one.iterations);
        std::printf("%-38s %9.4f %9.4f %9.4f\n", one.name, nearest, bilinear, bicubic);
    }

    std::printf("\nmilliseconds per resample, mean of the iterations shown in the source.\n");
    std::printf("Scaling only: cached-image creation and the Office apply are measured\n");
    std::printf("separately by the in-PowerPoint harnesses.\n");
    return 0;
}
