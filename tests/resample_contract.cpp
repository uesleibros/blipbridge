/**
 * @file resample_contract.cpp
 * Correctness tests for the BGRA32 resampling filters.
 *
 * Pure image maths - no Office, no COM - so this runs in CI on both
 * architectures and is the suite that has to catch a filter regression, because
 * nothing downstream would.
 *
 * Nearest is checked against exact expected bytes: it is defined as point
 * sampling, so "close enough" is not an acceptable answer for it. The
 * interpolating filters are checked against properties that pin the behaviour
 * without hard-coding float noise - identity, monotonicity, interpolation at
 * source points, edge clamping, and the alpha handling that stops dark halos.
 */

#include "../src/image/resample.hpp"

#include <array>
#include <cstdio>
#include <cstdlib>
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

using bb::image::Resample;
using bb::image::ResampleStatus;
using bb::image::ScaleFilter;

/// A BGRA image built from per-pixel values, tightly packed unless padded.
struct Image {
    std::uint32_t width = 0;
    std::uint32_t height = 0;
    std::uint32_t stride = 0;
    std::vector<std::uint8_t> pixels;

    const std::uint8_t* at(std::uint32_t x, std::uint32_t y) const {
        return pixels.data() + static_cast<std::size_t>(y) * stride + x * 4u;
    }
};

Image MakeImage(std::uint32_t width,
                std::uint32_t height,
                const std::vector<std::array<std::uint8_t, 4>>& values,
                std::uint32_t padding = 0) {
    Image image;
    image.width = width;
    image.height = height;
    image.stride = width * 4u + padding;
    // 0xCD in the padding: if a filter ever reads past a row, the result will be
    // visibly wrong rather than plausibly wrong.
    image.pixels.assign(static_cast<std::size_t>(image.stride) * height, 0xCD);
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            const auto& value = values[static_cast<std::size_t>(y) * width + x];
            std::uint8_t* pixel =
                image.pixels.data() + static_cast<std::size_t>(y) * image.stride + x * 4u;
            pixel[0] = value[0];
            pixel[1] = value[1];
            pixel[2] = value[2];
            pixel[3] = value[3];
        }
    }
    return image;
}

std::string Describe(const std::uint8_t* pixel) {
    char buffer[32];
    std::snprintf(buffer, sizeof(buffer), "%02X%02X%02X%02X", pixel[0], pixel[1], pixel[2],
                  pixel[3]);
    return buffer;
}

bool SamePixel(const std::uint8_t* a, const std::uint8_t* b) {
    return a[0] == b[0] && a[1] == b[1] && a[2] == b[2] && a[3] == b[3];
}

} // namespace

int main() {
    using Value = std::array<std::uint8_t, 4>;

    // Four distinguishable opaque pixels: A B / C D.
    const Value A{10, 20, 30, 255};
    const Value B{200, 40, 60, 255};
    const Value C{70, 210, 90, 255};
    const Value D{110, 130, 240, 255};
    const Image quad = MakeImage(2, 2, {A, B, C, D});

    std::vector<std::uint8_t> out;

    // --- nearest is exact ----------------------------------------------------
    {
        const ResampleStatus status = Resample(quad.pixels.data(), 2, 2, quad.stride, 4, 4,
                                               ScaleFilter::Nearest, out);
        Check(status == ResampleStatus::Ok, "nearest 2x2 -> 4x4 succeeds");
        // Doubling must replicate blocks exactly:
        //     A A B B
        //     A A B B
        //     C C D D
        //     C C D D
        const Value* expected[4][4] = {{&A, &A, &B, &B},
                                       {&A, &A, &B, &B},
                                       {&C, &C, &D, &D},
                                       {&C, &C, &D, &D}};
        bool exact = true;
        for (std::uint32_t y = 0; y < 4 && exact; ++y) {
            for (std::uint32_t x = 0; x < 4; ++x) {
                const std::uint8_t* got = out.data() + (static_cast<std::size_t>(y) * 4 + x) * 4;
                if (!SamePixel(got, expected[y][x]->data())) {
                    std::printf("       at (%u,%u) got %s\n", x, y, Describe(got).c_str());
                    exact = false;
                    break;
                }
            }
        }
        Check(exact, "nearest 2x doubling replicates blocks exactly");
    }

    // --- nearest preserves the source bytes, alpha included ------------------
    {
        const Value translucent{1, 2, 3, 7};
        const Image single = MakeImage(1, 1, {translucent});
        Resample(single.pixels.data(), 1, 1, single.stride, 5, 3, ScaleFilter::Nearest, out);
        bool same = true;
        for (std::size_t i = 0; i < out.size(); i += 4) {
            if (!SamePixel(out.data() + i, translucent.data())) {
                same = false;
                break;
            }
        }
        Check(same, "nearest 1x1 -> 5x3 copies the source pixel exactly, alpha included");
    }

    // --- odd, non-square, non-integer ratios ---------------------------------
    {
        Check(Resample(quad.pixels.data(), 2, 2, quad.stride, 3, 7, ScaleFilter::Nearest, out) ==
                  ResampleStatus::Ok,
              "nearest handles an odd non-square target");
        Check(out.size() == 3u * 7u * 4u, "and produces exactly the requested bytes");
        Check(Resample(quad.pixels.data(), 2, 2, quad.stride, 5, 3, ScaleFilter::Bilinear, out) ==
                  ResampleStatus::Ok,
              "bilinear handles a non-integer ratio");
        Check(Resample(quad.pixels.data(), 2, 2, quad.stride, 5, 3, ScaleFilter::Bicubic, out) ==
                  ResampleStatus::Ok,
              "bicubic handles a non-integer ratio");
    }

    // --- downscaling ---------------------------------------------------------
    {
        std::vector<Value> big(8 * 8, Value{80, 90, 100, 255});
        const Image source = MakeImage(8, 8, big);
        for (const ScaleFilter filter :
             {ScaleFilter::Nearest, ScaleFilter::Bilinear, ScaleFilter::Bicubic}) {
            const ResampleStatus status =
                Resample(source.pixels.data(), 8, 8, source.stride, 3, 3, filter, out);
            bool flat = status == ResampleStatus::Ok;
            for (std::size_t i = 0; flat && i < out.size(); i += 4) {
                // A flat field must stay flat under every filter; a normalisation
                // or premultiply error shows up here as drift.
                if (out[i] != 80 || out[i + 1] != 90 || out[i + 2] != 100 || out[i + 3] != 255) {
                    flat = false;
                }
            }
            Check(flat, std::string("downscaling a flat field is exact with ") +
                            bb::image::FilterName(filter));
        }
    }

    // --- identity: same size means untouched bytes ---------------------------
    {
        for (const ScaleFilter filter :
             {ScaleFilter::Nearest, ScaleFilter::Bilinear, ScaleFilter::Bicubic}) {
            Resample(quad.pixels.data(), 2, 2, quad.stride, 2, 2, filter, out);
            const bool same = SamePixel(out.data() + 0, A.data()) &&
                              SamePixel(out.data() + 4, B.data()) &&
                              SamePixel(out.data() + 8, C.data()) &&
                              SamePixel(out.data() + 12, D.data());
            Check(same, std::string("same-size resample is bit-exact with ") +
                            bb::image::FilterName(filter));
        }
    }

    // --- padded stride is honoured -------------------------------------------
    {
        const Image padded = MakeImage(2, 2, {A, B, C, D}, 11);
        Resample(padded.pixels.data(), 2, 2, padded.stride, 4, 4, ScaleFilter::Nearest, out);
        const bool corners = SamePixel(out.data(), A.data()) &&
                             SamePixel(out.data() + 3 * 4, B.data()) &&
                             SamePixel(out.data() + (3 * 4 + 0) * 4, C.data()) &&
                             SamePixel(out.data() + (3 * 4 + 3) * 4, D.data());
        Check(corners, "a padded source stride is honoured, not treated as pixels");
    }

    // --- alpha: no dark halo at a transparent edge ---------------------------
    {
        /*
         * The classic artefact. A transparent pixel whose colour bytes are black
         * sits beside an opaque white one. Interpolating straight BGRA would
         * average the black in and darken the edge; premultiplied interpolation
         * must not, because a zero-alpha pixel contributes no colour at all.
         */
        const Value clearBlack{0, 0, 0, 0};
        const Value opaqueWhite{255, 255, 255, 255};
        const Image edge = MakeImage(2, 1, {clearBlack, opaqueWhite});

        Resample(edge.pixels.data(), 2, 1, edge.stride, 8, 1, ScaleFilter::Bilinear, out);
        bool clean = true;
        for (std::uint32_t x = 0; x < 8; ++x) {
            const std::uint8_t* pixel = out.data() + static_cast<std::size_t>(x) * 4;
            // Wherever anything is visible at all, it must still be white.
            if (pixel[3] > 8 && (pixel[0] < 250 || pixel[1] < 250 || pixel[2] < 250)) {
                std::printf("       halo at x=%u: %s\n", x, Describe(pixel).c_str());
                clean = false;
            }
        }
        Check(clean, "bilinear does not darken a transparent edge");

        Resample(edge.pixels.data(), 2, 1, edge.stride, 8, 1, ScaleFilter::Bicubic, out);
        clean = true;
        for (std::uint32_t x = 0; x < 8; ++x) {
            const std::uint8_t* pixel = out.data() + static_cast<std::size_t>(x) * 4;
            if (pixel[3] > 8 && (pixel[0] < 240 || pixel[1] < 240 || pixel[2] < 240)) {
                std::printf("       halo at x=%u: %s\n", x, Describe(pixel).c_str());
                clean = false;
            }
        }
        Check(clean, "bicubic does not darken a transparent edge");
    }

    // --- interpolating filters pass through their source samples -------------
    {
        /*
         * Bilinear and Catmull-Rom both interpolate, so an exact 3x upscale
         * samples source centres at destination 1, 4, 7... and must reproduce
         * them. This is what pins the pixel-centre convention: a half-pixel
         * error moves these and the check fails.
         */
        const Image line = MakeImage(3, 1, {A, B, C});
        for (const ScaleFilter filter : {ScaleFilter::Bilinear, ScaleFilter::Bicubic}) {
            Resample(line.pixels.data(), 3, 1, line.stride, 9, 1, filter, out);
            const bool centres = SamePixel(out.data() + 1 * 4, A.data()) &&
                                 SamePixel(out.data() + 4 * 4, B.data()) &&
                                 SamePixel(out.data() + 7 * 4, C.data());
            Check(centres, std::string("source samples survive a 3x upscale with ") +
                               bb::image::FilterName(filter));
        }
    }

    // --- rejections ----------------------------------------------------------
    {
        Check(Resample(nullptr, 2, 2, 8, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::NullSource,
              "a null source is rejected");
        Check(Resample(quad.pixels.data(), 0, 2, 8, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidSourceSize,
              "a zero source width is rejected");
        Check(Resample(quad.pixels.data(), 2, 0, 8, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidSourceSize,
              "a zero source height is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, 8, 0, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidTargetSize,
              "a zero target width is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, 8, 4, 0, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidTargetSize,
              "a zero target height is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, 7, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidStride,
              "a stride below width*4 is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, 0, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidStride,
              "a zero stride is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, -8, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::InvalidStride,
              "a negative stride is rejected");
        Check(Resample(quad.pixels.data(), 2, 2, 8, 4, 4, static_cast<ScaleFilter>(99), out) ==
                  ResampleStatus::UnknownFilter,
              "an unknown filter is rejected");

        // Sizes whose product overflows 32 bits must be refused by the limit,
        // not turned into a short allocation and a long read.
        Check(Resample(quad.pixels.data(), 2, 2, 8, 100000, 100000, ScaleFilter::Nearest, out) ==
                  ResampleStatus::TooLarge,
              "an oversized target is rejected before allocating");
        Check(Resample(quad.pixels.data(), 65536, 65536, 8, 4, 4, ScaleFilter::Nearest, out) ==
                  ResampleStatus::TooLarge,
              "an oversized source is rejected");
    }

    // --- the filter table agrees with itself ---------------------------------
    {
        Check(bb::image::IsSupportedFilter(0) && bb::image::IsSupportedFilter(1) &&
                  bb::image::IsSupportedFilter(2),
              "every named filter reports as supported");
        Check(!bb::image::IsSupportedFilter(3) && !bb::image::IsSupportedFilter(0xFFFFFFFFu),
              "and nothing else does - no filter is named before it works");
    }

    if (g_failures != 0) {
        std::printf("\n%d resample checks failed\n", g_failures);
        return 1;
    }
    std::printf("\nall resample checks passed\n");
    return 0;
}
