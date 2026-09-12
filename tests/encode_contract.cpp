/**
 * @file encode_contract.cpp
 * Correctness tests for PNG encoding, and for the encode/decode round trip.
 *
 * No Office and no COM beyond WIC itself, so this runs in CI on both
 * architectures - which matters more here than for most of `src/image`, because
 * encoding exists *for* the portable backend and the portable backend is what
 * makes 32-bit PowerPoint work at all. Every texture that backend applies goes
 * through this code.
 *
 * ## What is asserted, and why it can be asserted exactly
 *
 * PNG is lossless, so the round trip is pixel-exact and is checked byte for
 * byte - including the alpha channel, which is the part most likely to be
 * quietly wrong. Straight alpha in, straight alpha out: if anything in the chain
 * premultiplied, a half-transparent red would come back darker, and the
 * transparency fixture would catch it.
 *
 * The refusals are asserted too. An encoder that accepts a stride narrower than
 * a row does not fail - it reads into the next row, and produces an image that
 * is subtly sheared rather than an error, which is the kind of defect that
 * survives a long time.
 */

#include "../src/image/decode.hpp"
#include "../src/image/encode.hpp"
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
using bb::image::DecodeStatus;
using bb::image::EncodePng;
using bb::image::EncodeStatus;

struct Bgra {
    std::uint8_t b, g, r, a;
};

/**
 * A pattern where every pixel is a function of its position.
 *
 * That is the point: a uniform image would round-trip successfully even if rows
 * or columns were transposed, and a corner-coloured one would survive a small
 * offset. Here every pixel is distinct, so any displacement at all shows up.
 */
std::vector<std::uint8_t> PositionPattern(std::uint32_t width, std::uint32_t height, bool opaque) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(width) * height * 4);
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            Bgra colour{};
            colour.b = static_cast<std::uint8_t>(x * 7 + 3);
            colour.g = static_cast<std::uint8_t>(y * 11 + 5);
            colour.r = static_cast<std::uint8_t>((x + y) * 13 + 17);
            colour.a = opaque ? 255 : static_cast<std::uint8_t>(x * 9 + y * 3);
            std::memcpy(&pixels[(static_cast<std::size_t>(y) * width + x) * 4], &colour, 4);
        }
    }
    return pixels;
}

/// Encodes, decodes, and asserts the pixels came back exactly as they went in.
void RoundTrip(const char* name, std::uint32_t width, std::uint32_t height, bool opaque) {
    const std::vector<std::uint8_t> source = PositionPattern(width, height, opaque);

    std::vector<std::uint8_t> encoded;
    const EncodeStatus status =
        EncodePng(source.data(), width, height, static_cast<std::int32_t>(width) * 4, encoded);
    Check(status == EncodeStatus::Ok, std::string(name) + ": encoded");
    if (status != EncodeStatus::Ok) {
        return;
    }

    // A PNG file starts with a fixed eight-byte signature. Checking it proves
    // the container is what was asked for, not merely that bytes came back.
    static const std::uint8_t kPngSignature[8] = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A};
    Check(encoded.size() > 8 && std::memcmp(encoded.data(), kPngSignature, 8) == 0,
          std::string(name) + ": the bytes are a PNG");

    DecodedImage decoded;
    const DecodeStatus back = bb::image::Decode(encoded.data(), encoded.size(), decoded);
    Check(back == DecodeStatus::Ok, std::string(name) + ": decoded again");
    if (back != DecodeStatus::Ok) {
        return;
    }
    Check(decoded.width == width && decoded.height == height,
          std::string(name) + ": dimensions survived");
    Check(decoded.pixels.size() == source.size(), std::string(name) + ": byte count survived");
    Check(decoded.pixels == source, std::string(name) + ": every pixel survived exactly");
}

/// The round trip with a padded source, which is what a cropped buffer looks like.
void RoundTripWithStride() {
    constexpr std::uint32_t width = 13;
    constexpr std::uint32_t height = 7;
    constexpr std::int32_t stride = width * 4 + 24; // deliberately not a tidy number

    std::vector<std::uint8_t> padded(static_cast<std::size_t>(stride) * height, 0xCD);
    const std::vector<std::uint8_t> tight = PositionPattern(width, height, false);
    for (std::uint32_t y = 0; y < height; ++y) {
        std::memcpy(&padded[static_cast<std::size_t>(stride) * y],
                    &tight[static_cast<std::size_t>(width) * 4 * y],
                    static_cast<std::size_t>(width) * 4);
    }

    std::vector<std::uint8_t> encoded;
    Check(EncodePng(padded.data(), width, height, stride, encoded) == EncodeStatus::Ok,
          "padded: encoded");

    DecodedImage decoded;
    Check(bb::image::Decode(encoded.data(), encoded.size(), decoded) == DecodeStatus::Ok,
          "padded: decoded again");
    // The padding bytes must not have reached the image: if the encoder ignored
    // the stride, the filler would appear in the pixels.
    Check(decoded.pixels == tight, "padded: the row padding stayed out of the image");
}

void Refusals() {
    std::vector<std::uint8_t> out;
    const std::vector<std::uint8_t> pixels = PositionPattern(4, 4, true);

    Check(EncodePng(nullptr, 4, 4, 16, out) == EncodeStatus::InvalidArgument,
          "refuses a null pixel pointer");
    Check(EncodePng(pixels.data(), 0, 4, 16, out) == EncodeStatus::InvalidArgument,
          "refuses a zero width");
    Check(EncodePng(pixels.data(), 4, 0, 16, out) == EncodeStatus::InvalidArgument,
          "refuses a zero height");
    Check(EncodePng(pixels.data(), 4, 4, 15, out) == EncodeStatus::InvalidArgument,
          "refuses a stride narrower than one row");
    // 64 megapixels is the ceiling the resampler enforces; the encoder must
    // agree, or a pipeline could produce something it cannot then write out.
    Check(EncodePng(pixels.data(), 65536, 16384, 65536 * 4, out) == EncodeStatus::UnsupportedSize,
          "refuses more pixels than the library's ceiling");

    // A refused encode must leave the caller's buffer alone rather than half
    // filling it, so a caller who ignores the status is not silently handed
    // fragments of a previous image.
    out.assign(4, 0x42);
    Check(EncodePng(nullptr, 4, 4, 16, out) == EncodeStatus::InvalidArgument && out.size() == 4 &&
              out[0] == 0x42,
          "a refusal leaves the output buffer untouched");
}

/// Every status must describe itself; an unnamed one reaches a user as nothing.
void Descriptions() {
    const EncodeStatus all[] = {EncodeStatus::Ok,
                                EncodeStatus::InvalidArgument,
                                EncodeStatus::NoEncoder,
                                EncodeStatus::UnsupportedSize,
                                EncodeStatus::EncodeFailed,
                                EncodeStatus::OutOfMemory};
    bool described = true;
    for (const EncodeStatus status : all) {
        const char* text = bb::image::DescribeEncodeStatus(status);
        described = described && text != nullptr && *text != '\0';
    }
    Check(described, "every status has a message");
}

} // namespace

int main() {
    std::printf("encode contract\n");
    RoundTrip("opaque", 16, 9, true);
    RoundTrip("transparent", 16, 9, false);
    RoundTrip("single pixel", 1, 1, false);
    RoundTrip("tall", 3, 64, true);
    RoundTripWithStride();
    Refusals();
    Descriptions();

    if (g_failures != 0) {
        std::printf("%d failure(s)\n", g_failures);
        return 1;
    }
    std::printf("all checks passed\n");
    return 0;
}
