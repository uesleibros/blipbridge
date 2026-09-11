/**
 * @file image_benchmark.cpp
 * Timings for the image pipeline: decode, scale, crop, warp.
 *
 * Built so it cannot rot, never registered as a test: a busy machine must not
 * turn a measurement into a failure. Run it by hand.
 *
 * Every figure is from one machine in one session. What is worth reading is the
 * *shape* - which stage dominates, how cost grows with size, whether something
 * fits a frame budget - not the absolute milliseconds, which move by a factor of
 * two or more with whatever else is running.
 */

#include "../src/image/pipeline.hpp"
#include "../src/image/warp.hpp"

#include <windows.h>

#include <algorithm>
#include <cstdio>
#include <string>
#include <vector>
#include <wincodec.h>

namespace {

double SecondsPerTick() {
    LARGE_INTEGER frequency{};
    QueryPerformanceFrequency(&frequency);
    return 1.0 / static_cast<double>(frequency.QuadPart);
}

long long Now() {
    LARGE_INTEGER counter{};
    QueryPerformanceCounter(&counter);
    return counter.QuadPart;
}

/// Median of repeated runs, which is what survives a noisy machine.
template <typename Body>
double MedianMs(int rounds, Body body) {
    const double tick = SecondsPerTick();
    std::vector<double> samples;
    samples.reserve(rounds);
    for (int round = 0; round < rounds; ++round) {
        const long long start = Now();
        body();
        samples.push_back(static_cast<double>(Now() - start) * tick * 1000.0);
    }
    std::sort(samples.begin(), samples.end());
    return samples[samples.size() / 2];
}

std::vector<std::uint8_t> Photo(std::uint32_t size) {
    // Enough structure that a codec cannot trivially compress it away.
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(size) * size * 4);
    for (std::uint32_t y = 0; y < size; ++y) {
        for (std::uint32_t x = 0; x < size; ++x) {
            auto* p = &pixels[(static_cast<std::size_t>(y) * size + x) * 4];
            p[0] = static_cast<std::uint8_t>((x * 7 + y * 3) & 0xFF);
            p[1] = static_cast<std::uint8_t>((x ^ y) & 0xFF);
            p[2] = static_cast<std::uint8_t>((x * y) & 0xFF);
            p[3] = 255;
        }
    }
    return pixels;
}

bool Encode(const std::vector<std::uint8_t>& pixels,
            std::uint32_t size,
            const GUID& container,
            std::vector<std::uint8_t>& out) {
    IWICImagingFactory* factory = nullptr;
    HRESULT hr = CoCreateInstance(
        CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
    if (hr == CO_E_NOTINITIALIZED) {
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        hr = CoCreateInstance(
            CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
    }
    if (FAILED(hr)) {
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
        GUID format = GUID_WICPixelFormat32bppBGRA;
        if (SUCCEEDED(frame->SetSize(size, size)) && SUCCEEDED(frame->SetPixelFormat(&format)) &&
            SUCCEEDED(frame->WritePixels(size,
                                         size * 4,
                                         static_cast<UINT>(pixels.size()),
                                         const_cast<BYTE*>(pixels.data()))) &&
            SUCCEEDED(frame->Commit()) && SUCCEEDED(encoder->Commit())) {
            HGLOBAL memory = nullptr;
            STATSTG stat{};
            if (SUCCEEDED(GetHGlobalFromStream(stream, &memory)) &&
                SUCCEEDED(stream->Stat(&stat, STATFLAG_NONAME))) {
                const auto* bytes = static_cast<const std::uint8_t*>(GlobalLock(memory));
                if (bytes) {
                    out.assign(bytes, bytes + static_cast<std::size_t>(stat.cbSize.QuadPart));
                    ok = !out.empty();
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

const char* FilterLabel(bb::image::ScaleFilter filter) {
    return bb::image::FilterName(filter);
}

} // namespace

int main() {
    using namespace bb::image;

    std::printf("image pipeline benchmark - one machine, one session, medians of repeated runs\n");
    std::printf("absolute times move with machine load; the shape of the numbers is the point\n\n");

    for (const std::uint32_t size : {256u, 1024u}) {
        const auto source = Photo(size);
        std::vector<std::uint8_t> png;
        std::vector<std::uint8_t> jpeg;
        if (!Encode(source, size, GUID_ContainerFormatPng, png) ||
            !Encode(source, size, GUID_ContainerFormatJpeg, jpeg)) {
            std::printf("could not build the %ux%u fixtures\n", size, size);
            continue;
        }

        std::printf("== %ux%u source (PNG %zu KB, JPEG %zu KB) ==\n",
                    size, size, png.size() / 1024, jpeg.size() / 1024);

        const int rounds = size >= 1024 ? 20 : 60;

        DecodedImage decoded;
        const double pngDecode =
            MedianMs(rounds, [&] { Decode(png.data(), png.size(), decoded); });
        const double jpegDecode =
            MedianMs(rounds, [&] { Decode(jpeg.data(), jpeg.size(), decoded); });
        std::printf("  decode PNG                       %8.3f ms\n", pngDecode);
        std::printf("  decode JPEG                      %8.3f ms\n", jpegDecode);

        // Decode plus scale in one call, which is what the new path does, against
        // the decode alone - so the resample's share is visible rather than
        // implied.
        for (const ScaleFilter filter :
             {ScaleFilter::Nearest, ScaleFilter::Bilinear, ScaleFilter::Bicubic}) {
            PrepareRequest request;
            request.targetWidth = size / 2;
            request.targetHeight = size / 2;
            request.filter = filter;
            const double total =
                MedianMs(rounds, [&] { PrepareEncoded(png.data(), png.size(), request); });
            std::printf("  PNG -> half size, %-9s      %8.3f ms  (decode %.3f + scale %.3f)\n",
                        FilterLabel(filter), total, pngDecode, total - pngDecode);
        }

        // The atlas case: one tile out of the image, scaled up.
        {
            PrepareRequest request;
            request.cropX = 0;
            request.cropY = 0;
            request.cropWidth = 16;
            request.cropHeight = 16;
            request.targetWidth = 256;
            request.targetHeight = 256;
            request.filter = ScaleFilter::Nearest;
            const double total =
                MedianMs(rounds, [&] { PrepareEncoded(png.data(), png.size(), request); });
            std::printf("  PNG -> 16x16 tile -> 256x256     %8.3f ms  (decode %.3f + crop/scale %.3f)\n",
                        total, pngDecode, total - pngDecode);
        }

        // Crop and scale from pixels already in hand, which is what a caller
        // pulling many tiles out of one decoded atlas actually pays per tile.
        {
            PrepareRequest request;
            request.cropWidth = 16;
            request.cropHeight = 16;
            request.targetWidth = 256;
            request.targetHeight = 256;
            request.filter = ScaleFilter::Nearest;
            const double perTile = MedianMs(rounds * 4, [&] {
                PreparePixels(source.data(), size, size,
                              static_cast<std::int32_t>(size) * 4, request);
            });
            std::printf("  one more tile from that decode   %8.3f ms\n", perTile);
        }
        std::printf("\n");
    }

    // --- the quad warp ------------------------------------------------------
    std::printf("== projective warp, from a 256x256 source onto a tapered quad ==\n");
    const auto source = Photo(256);
    for (const std::uint32_t output : {16u, 64u, 256u, 512u, 1024u}) {
        const PointF quad[4] = {{static_cast<float>(output) * 0.35f, 0.0f},
                                {static_cast<float>(output) * 0.65f, 0.0f},
                                {static_cast<float>(output), static_cast<float>(output)},
                                {0.0f, static_cast<float>(output)}};
        for (const ScaleFilter filter : {ScaleFilter::Nearest, ScaleFilter::Bilinear}) {
            const int rounds = output >= 512 ? 20 : 200;
            const double ms = MedianMs(rounds, [&] {
                WarpQuad(source.data(), 256, 256, 256 * 4, quad, output, output, filter);
            });
            std::printf("  %4ux%-4u %-9s %8.3f ms   %s\n",
                        output, output, FilterLabel(filter), ms,
                        ms <= 16.67 ? "fits a 60 FPS frame on its own" : "does not fit 16.67 ms");
        }
    }

    std::printf("\nA frame budget note: \"fits\" above means the warp alone. A real frame also\n"
                "pays for creating the image resource and for Office's own apply, which the\n"
                "apply benchmarks measure separately.\n");
    return 0;
}
