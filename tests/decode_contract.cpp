/**
 * @file decode_contract.cpp
 * Correctness tests for encoded-image decoding, and for decode feeding the
 * resampler.
 *
 * No Office and no COM beyond WIC itself, so this runs in CI on both
 * architectures - which matters, because the decoder is the first piece of
 * BlipBridge that is genuinely useful on x86 even though the Office backend
 * there refuses everything.
 *
 * ## Fixtures are generated, not checked in
 *
 * Every image here is encoded at run time by WIC from a pattern this file
 * builds, so the tests carry no binary blobs and cannot drift from what they
 * claim to contain. PNG and BMP are lossless, so those round-trip pixel-exact
 * and are asserted byte for byte. JPEG is not, so it is asserted on size and on
 * colours being recognisably themselves rather than on exact bytes - claiming
 * more would be claiming something false about JPEG.
 */

#include "../src/image/decode.hpp"
#include "../src/image/resample.hpp"

#include <wincodec.h>

#include <algorithm>
#include <cstdio>
#include <cstdlib>
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
using bb::image::Resample;
using bb::image::ResampleStatus;
using bb::image::ScaleFilter;

struct Bgra {
    std::uint8_t b, g, r, a;
};

/// The four-quadrant fixture: corner colours, so orientation errors are visible.
std::vector<std::uint8_t> CornerPattern(std::uint32_t width, std::uint32_t height, bool opaque) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(width) * height * 4);
    for (std::uint32_t y = 0; y < height; ++y) {
        for (std::uint32_t x = 0; x < width; ++x) {
            const bool right = x >= width / 2;
            const bool bottom = y >= height / 2;
            // top-left red, top-right green, bottom-right blue, bottom-left yellow
            Bgra colour{};
            if (!right && !bottom) {
                colour = {0, 0, 255, 255};
            } else if (right && !bottom) {
                colour = {0, 255, 0, 255};
            } else if (right && bottom) {
                colour = {255, 0, 0, 255};
            } else {
                colour = {0, 255, 255, 255};
            }
            if (!opaque && !right && bottom) {
                // One quadrant fully transparent, so alpha survival is testable.
                colour = {0, 0, 0, 0};
            }
            auto* out = &pixels[(static_cast<std::size_t>(y) * width + x) * 4];
            out[0] = colour.b;
            out[1] = colour.g;
            out[2] = colour.r;
            out[3] = colour.a;
        }
    }
    return pixels;
}

/**
 * Encodes BGRA32 with WIC, so the fixtures are real files of that format.
 *
 * The apartment fallback matters: this suite deliberately does not initialise
 * COM in main, because the decoder is documented to cope with an uninitialised
 * apartment and that is worth testing rather than arranging away. The fixture
 * builder has to cope with the same thing or it fails for a reason that has
 * nothing to do with the code under test - which is exactly what it did.
 */
bool Encode(const std::vector<std::uint8_t>& pixels,
            std::uint32_t width,
            std::uint32_t height,
            const GUID& container,
            bool withAlpha,
            std::vector<std::uint8_t>& out) {
    IWICImagingFactory* factory = nullptr;
    HRESULT created = CoCreateInstance(
        CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
    if (created == CO_E_NOTINITIALIZED) {
        if (SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
            created = CoCreateInstance(
                CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
        }
    }
    if (FAILED(created)) {
        return false;
    }

    bool ok = false;
    IStream* stream = nullptr;
    IWICBitmapEncoder* encoder = nullptr;
    IWICBitmapFrameEncode* frame = nullptr;
    IPropertyBag2* options = nullptr;
    if (SUCCEEDED(CreateStreamOnHGlobal(nullptr, TRUE, &stream)) &&
        SUCCEEDED(factory->CreateEncoder(container, nullptr, &encoder)) &&
        SUCCEEDED(encoder->Initialize(stream, WICBitmapEncoderNoCache)) &&
        SUCCEEDED(encoder->CreateNewFrame(&frame, &options)) &&
        SUCCEEDED(frame->Initialize(options))) {
        GUID format = withAlpha ? GUID_WICPixelFormat32bppBGRA : GUID_WICPixelFormat24bppBGR;

        // A 24bpp frame is fed 24bpp rows. Handing it the 32bpp buffer would
        // encode three of every four bytes as a pixel and shear the image - and
        // the failure looks like a decoder bug rather than a fixture one.
        std::vector<std::uint8_t> packed;
        const std::uint8_t* rows = pixels.data();
        UINT stride = width * 4;
        if (!withAlpha) {
            packed.resize(static_cast<std::size_t>(width) * height * 3);
            for (std::size_t index = 0, count = static_cast<std::size_t>(width) * height;
                 index < count;
                 ++index) {
                packed[index * 3 + 0] = pixels[index * 4 + 0];
                packed[index * 3 + 1] = pixels[index * 4 + 1];
                packed[index * 3 + 2] = pixels[index * 4 + 2];
            }
            rows = packed.data();
            stride = width * 3;
        }
        const auto payload = static_cast<UINT>(withAlpha ? pixels.size() : packed.size());

        if (SUCCEEDED(frame->SetSize(width, height)) &&
            SUCCEEDED(frame->SetPixelFormat(&format)) &&
            SUCCEEDED(frame->WritePixels(height, stride, payload, const_cast<BYTE*>(rows))) &&
            SUCCEEDED(frame->Commit()) && SUCCEEDED(encoder->Commit())) {
            HGLOBAL memory = nullptr;
            if (SUCCEEDED(GetHGlobalFromStream(stream, &memory))) {
                const auto size = static_cast<std::size_t>(GlobalSize(memory));
                const auto* bytes = static_cast<const std::uint8_t*>(GlobalLock(memory));
                if (bytes) {
                    // GlobalSize rounds up; the stream's own position is the
                    // length that was actually written.
                    STATSTG stat{};
                    if (SUCCEEDED(stream->Stat(&stat, STATFLAG_NONAME))) {
                        const auto written = static_cast<std::size_t>(stat.cbSize.QuadPart);
                        out.assign(bytes, bytes + (written <= size ? written : size));
                        ok = !out.empty();
                    }
                    GlobalUnlock(memory);
                }
            }
        }
    }
    if (options) options->Release();
    if (frame) frame->Release();
    if (encoder) encoder->Release();
    if (stream) stream->Release();
    factory->Release();
    return ok;
}

/// True when every byte matches - the only acceptable answer for a lossless format.
bool SamePixels(const std::vector<std::uint8_t>& a, const std::vector<std::uint8_t>& b) {
    return a.size() == b.size() && std::equal(a.begin(), a.end(), b.begin());
}

Bgra At(const DecodedImage& image, std::uint32_t x, std::uint32_t y) {
    const auto* p = &image.pixels[(static_cast<std::size_t>(y) * image.width + x) * 4];
    return Bgra{p[0], p[1], p[2], p[3]};
}

/// Loose comparison, for a format that does not promise exactness.
bool Near(std::uint8_t actual, std::uint8_t expected, int tolerance) {
    return std::abs(static_cast<int>(actual) - static_cast<int>(expected)) <= tolerance;
}

void TestLosslessRoundTrip(const char* name, const GUID& container, bool withAlpha) {
    constexpr std::uint32_t kWidth = 8;
    constexpr std::uint32_t kHeight = 8;
    const std::vector<std::uint8_t> source = CornerPattern(kWidth, kHeight, !withAlpha);

    std::vector<std::uint8_t> encoded;
    if (!Encode(source, kWidth, kHeight, container, withAlpha, encoded)) {
        Check(false, std::string(name) + ": the fixture could be encoded");
        return;
    }

    DecodedImage decoded;
    const DecodeStatus status = bb::image::Decode(encoded.data(), encoded.size(), decoded);
    Check(status == DecodeStatus::Ok, std::string(name) + ": decodes");
    if (status != DecodeStatus::Ok) {
        return;
    }
    Check(decoded.width == kWidth && decoded.height == kHeight,
          std::string(name) + ": keeps its dimensions");
    Check(decoded.stride() == static_cast<std::int32_t>(kWidth) * 4,
          std::string(name) + ": decodes to a tightly packed buffer");
    Check(SamePixels(decoded.pixels, source),
          std::string(name) + ": round-trips pixel-exact through a lossless format");

    if (withAlpha) {
        const Bgra transparent = At(decoded, 1, kHeight - 1);
        Check(transparent.a == 0, std::string(name) + ": preserves a fully transparent quadrant");
        const Bgra opaque = At(decoded, 1, 1);
        Check(opaque.a == 255, std::string(name) + ": and leaves opaque pixels opaque");
    } else {
        Check(At(decoded, 1, 1).a == 255,
              std::string(name) + ": a format without alpha decodes fully opaque");
    }
}

void TestJpeg() {
    constexpr std::uint32_t kWidth = 16;
    constexpr std::uint32_t kHeight = 16;
    const std::vector<std::uint8_t> source = CornerPattern(kWidth, kHeight, true);

    std::vector<std::uint8_t> encoded;
    if (!Encode(source, kWidth, kHeight, GUID_ContainerFormatJpeg, false, encoded)) {
        Check(false, "JPEG: the fixture could be encoded");
        return;
    }

    DecodedImage decoded;
    const DecodeStatus status = bb::image::Decode(encoded.data(), encoded.size(), decoded);
    Check(status == DecodeStatus::Ok, "JPEG: decodes");
    if (status != DecodeStatus::Ok) {
        return;
    }
    Check(decoded.width == kWidth && decoded.height == kHeight, "JPEG: keeps its dimensions");
    Check(At(decoded, 1, 1).a == 255, "JPEG: decodes fully opaque");

    // Not byte-exact - it is JPEG - but a red corner has to still be red, and a
    // tolerance this wide would still catch a channel swap or a flip.
    const Bgra topLeft = At(decoded, 2, 2);
    const Bgra topRight = At(decoded, kWidth - 3, 2);
    Check(Near(topLeft.r, 255, 40) && Near(topLeft.g, 0, 40) && Near(topLeft.b, 0, 40),
          "JPEG: the red corner decodes red, so the channel order survived");
    Check(Near(topRight.g, 255, 40) && Near(topRight.r, 0, 40) && Near(topRight.b, 0, 40),
          "JPEG: the green corner decodes green, so the image is not mirrored");
}

void TestRefusals() {
    DecodedImage decoded;

    Check(bb::image::Decode(nullptr, 0, decoded) == DecodeStatus::NoData,
          "no data is refused as NoData");
    const std::uint8_t byte = 0;
    Check(bb::image::Decode(&byte, 0, decoded) == DecodeStatus::NoData,
          "a zero length is refused as NoData");

    const std::uint8_t noise[] = {'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'};
    const DecodeStatus unknown = bb::image::Decode(noise, sizeof(noise), decoded);
    Check(unknown == DecodeStatus::UnknownFormat || unknown == DecodeStatus::Corrupt,
          "arbitrary bytes are refused rather than decoded");

    // A real PNG cut in half: the header is valid, the data is not there. This
    // is the case a decoder is most likely to read past, so it is the one worth
    // having.
    const std::vector<std::uint8_t> source = CornerPattern(8, 8, true);
    std::vector<std::uint8_t> encoded;
    if (Encode(source, 8, 8, GUID_ContainerFormatPng, true, encoded) && encoded.size() > 40) {
        const DecodeStatus truncated =
            bb::image::Decode(encoded.data(), encoded.size() / 2, decoded);
        Check(truncated != DecodeStatus::Ok, "a truncated PNG is refused, not half-decoded");

        std::vector<std::uint8_t> corrupted = encoded;
        for (std::size_t index = encoded.size() / 2; index < corrupted.size(); ++index) {
            corrupted[index] = static_cast<std::uint8_t>(index * 31);
        }
        const DecodeStatus damaged =
            bb::image::Decode(corrupted.data(), corrupted.size(), decoded);
        Check(damaged != DecodeStatus::Ok || decoded.width == 8,
              "a PNG with a damaged tail either fails or still reports its real size");

        const std::uint8_t headerOnly[] = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A};
        Check(bb::image::Decode(headerOnly, sizeof(headerOnly), decoded) != DecodeStatus::Ok,
              "a PNG signature with no image behind it is refused");
    } else {
        Check(false, "the truncation fixture could be built");
    }

    Check(bb::image::DecodeFile(nullptr, decoded) == DecodeStatus::NoData,
          "a null path is refused");
    Check(bb::image::DecodeFile(L"", decoded) == DecodeStatus::NoData,
          "an empty path is refused");
    Check(bb::image::DecodeFile(L"Z:\\no\\such\\file.png", decoded) == DecodeStatus::NoData,
          "a missing file is refused as NoData rather than as a format problem");
}

/// The decode output has to be exactly what the resampler documents as input.
void TestFeedsResampler() {
    constexpr std::uint32_t kWidth = 4;
    constexpr std::uint32_t kHeight = 4;
    const std::vector<std::uint8_t> source = CornerPattern(kWidth, kHeight, true);
    std::vector<std::uint8_t> encoded;
    if (!Encode(source, kWidth, kHeight, GUID_ContainerFormatPng, true, encoded)) {
        Check(false, "the resampler fixture could be built");
        return;
    }

    DecodedImage decoded;
    if (bb::image::Decode(encoded.data(), encoded.size(), decoded) != DecodeStatus::Ok) {
        Check(false, "the resampler fixture decodes");
        return;
    }

    // Identity: same size in and out, so the bytes must survive untouched
    // whichever filter is named.
    for (const ScaleFilter filter : {ScaleFilter::Nearest, ScaleFilter::Bilinear,
                                     ScaleFilter::Bicubic}) {
        std::vector<std::uint8_t> out;
        const ResampleStatus status = Resample(decoded.pixels.data(),
                                               decoded.width,
                                               decoded.height,
                                               decoded.stride(),
                                               decoded.width,
                                               decoded.height,
                                               filter,
                                               out);
        Check(status == ResampleStatus::Ok && SamePixels(out, decoded.pixels),
              std::string("decode then resample at the same size is bit-exact (") +
                  bb::image::FilterName(filter) + ")");
    }

    // Nearest doubling: every source pixel becomes a 2x2 block, exactly.
    std::vector<std::uint8_t> doubled;
    const ResampleStatus status = Resample(decoded.pixels.data(),
                                           decoded.width,
                                           decoded.height,
                                           decoded.stride(),
                                           kWidth * 2,
                                           kHeight * 2,
                                           ScaleFilter::Nearest,
                                           doubled);
    Check(status == ResampleStatus::Ok, "decode then nearest-double succeeds");
    if (status == ResampleStatus::Ok) {
        bool exact = true;
        for (std::uint32_t y = 0; y < kHeight * 2 && exact; ++y) {
            for (std::uint32_t x = 0; x < kWidth * 2; ++x) {
                const auto* got = &doubled[(static_cast<std::size_t>(y) * kWidth * 2 + x) * 4];
                const auto* want =
                    &decoded.pixels[(static_cast<std::size_t>(y / 2) * kWidth + x / 2) * 4];
                if (std::memcmp(got, want, 4) != 0) {
                    exact = false;
                    break;
                }
            }
        }
        Check(exact, "and every source pixel became an exact 2x2 block");
    }
}

/// Decoding the same bytes repeatedly must not leak or drift.
void TestRepeated() {
    const std::vector<std::uint8_t> source = CornerPattern(32, 32, true);
    std::vector<std::uint8_t> encoded;
    if (!Encode(source, 32, 32, GUID_ContainerFormatPng, true, encoded)) {
        Check(false, "the repetition fixture could be built");
        return;
    }

    bool stable = true;
    for (int round = 0; round < 200; ++round) {
        DecodedImage decoded;
        if (bb::image::Decode(encoded.data(), encoded.size(), decoded) != DecodeStatus::Ok ||
            !SamePixels(decoded.pixels, source)) {
            stable = false;
            break;
        }
    }
    Check(stable, "200 decodes of the same image all give the same pixels");
}

} // namespace

int main() {
    // The decoder creates its own apartment if it has to, which is exactly what
    // a VBA or Office caller will do to it. Left uninitialised here on purpose.
    std::printf("decode contract\n");

    TestLosslessRoundTrip("PNG RGBA", GUID_ContainerFormatPng, true);
    TestLosslessRoundTrip("PNG RGB", GUID_ContainerFormatPng, false);
    TestLosslessRoundTrip("BMP", GUID_ContainerFormatBmp, false);
    TestJpeg();
    TestRefusals();
    TestFeedsResampler();
    TestRepeated();

    if (g_failures != 0) {
        std::printf("%d check(s) failed\n", g_failures);
        return 1;
    }
    std::printf("all checks passed\n");
    return 0;
}
