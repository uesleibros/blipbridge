/**
 * @file texture_benchmark.cpp
 * In-process comparison of Fill.UserPicture against a reusable native texture.
 *
 * The measurement has to run inside PowerPoint. Driving it from PowerShell adds
 * a cross-process COM round trip per call, which is several times the cost of
 * the operation being measured and swamps the result; the earlier PowerShell
 * loops are stability checks, not benchmarks.
 *
 * Three quantities are reported separately, because they answer different
 * questions:
 *
 *  - `userPicture`  - one `Fill.UserPicture(path)` per Shape. Office re-decodes
 *                     on every call, which is the cost the texture handle is
 *                     meant to remove.
 *  - `loadTexture`  - decoding the bytes once into a cached image. Paid once per
 *                     texture, never on the hot path.
 *  - `applyTexture` - the hot path: record construction plus the receiver call,
 *                     against an already decoded image.
 *
 * All three fill the same Shapes with the same image in the same document, in a
 * fixed order, so the comparison is like for like. Timings come from
 * QueryPerformanceCounter around the call only.
 */

#include "../experiment_api.hpp"

#include <blipbridge/blipbridge.h>

#include "native_apply.hpp"
#include "oart_layout.hpp"

#include <blipbridge/dispatch.hpp>

#include <algorithm>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <vector>

namespace {

/// Enough repetitions to swamp scheduling noise without a long-running test.
constexpr int kDefaultIterations = 500;

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

/**
 * Per-call milliseconds. Percentiles matter more than the mean here: a fill that
 * is usually fast but occasionally stalls would look fine on the mean alone.
 */
struct Timing {
    double meanMs = 0.0;
    double medianMs = 0.0;
    double p95Ms = 0.0;
    double p99Ms = 0.0;
    double maxMs = 0.0;
    double totalMs = 0.0;
    int calls = 0;
};

/// Nearest-rank percentile on an already sorted sample.
double Percentile(const std::vector<double>& sorted, double fraction) {
    if (sorted.empty()) {
        return 0.0;
    }
    const std::size_t rank =
        static_cast<std::size_t>(fraction * static_cast<double>(sorted.size()));
    return sorted[std::min(rank, sorted.size() - 1)];
}

Timing Summarise(std::vector<double>& samples) {
    Timing timing;
    timing.calls = static_cast<int>(samples.size());
    if (samples.empty()) {
        return timing;
    }
    for (double sample : samples) {
        timing.totalMs += sample;
    }
    timing.meanMs = timing.totalMs / samples.size();
    std::sort(samples.begin(), samples.end());
    timing.medianMs = Percentile(samples, 0.50);
    timing.p95Ms = Percentile(samples, 0.95);
    timing.p99Ms = Percentile(samples, 0.99);
    timing.maxMs = samples.back();
    return timing;
}

void Append(std::wostringstream& out, const wchar_t* label, const Timing& timing) {
    out << label << L"Calls=" << timing.calls << L';' << label << L"MeanMs="
        << timing.meanMs << L';' << label << L"MedianMs=" << timing.medianMs << L';'
        << label << L"P95Ms=" << timing.p95Ms << L';' << label << L"P99Ms="
        << timing.p99Ms << L';' << label << L"MaxMs=" << timing.maxMs << L';' << label
        << L"TotalMs=" << timing.totalMs << L';';
}

/// Frees a SAFEARRAY built for one benchmark leg.
struct ArrayGuard {
    SAFEARRAY* value;
    ~ArrayGuard() { SafeArrayDestroy(value); }
};

} // namespace

/**
 * Benchmarks @p iterations fills several ways on the Shapes of @p slide.
 *
 * @p slide is a live PowerPoint Slide; the Shapes are created and removed here so
 * the caller's document is left as it was found. @p imagePath is used only by the
 * `UserPicture` leg - the native legs read the same file into memory once and
 * never touch the filesystem again.
 */
std::wstring benchmarkNativeTexture(IDispatch* slide, const std::wstring& imagePath,
                                    long iterations) {
    if (iterations <= 0) {
        iterations = kDefaultIterations;
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Benchmarks require the PowerPoint host");
    }

    std::ifstream file(std::filesystem::path(imagePath), std::ios::binary);
    if (!file) {
        throw bb::Error(E_INVALIDARG, "Cannot read the benchmark image");
    }
    const std::vector<char> contents((std::istreambuf_iterator<char>(file)),
                                     std::istreambuf_iterator<char>());
    if (contents.empty()) {
        throw bb::Error(E_INVALIDARG, "Benchmark image is empty");
    }

    SAFEARRAYBOUND bound{static_cast<ULONG>(contents.size()), 0};
    SAFEARRAY* bytes = SafeArrayCreate(VT_UI1, 1, &bound);
    if (!bytes) {
        throw std::bad_alloc();
    }
    ArrayGuard arrayGuard{bytes};
    void* raw = nullptr;
    bb::check(SafeArrayAccessData(bytes, &raw), "SafeArrayAccessData");
    std::memcpy(raw, contents.data(), contents.size());
    SafeArrayUnaccessData(bytes);

    // One Shape reused for every leg, so geometry and slide state are identical
    // across the comparison.
    bb::Value shapes = bb::get(slide, L"Shapes");
    bb::Value shape = bb::call(shapes.obj(), L"AddShape",
                               {bb::Value(1L), bb::Value(20.0), bb::Value(20.0),
                                bb::Value(120.0), bb::Value(120.0)});
    const double tick = SecondsPerTick();
    std::vector<double> userPicture;
    std::vector<double> applyTexture;
    userPicture.reserve(iterations);
    applyTexture.reserve(iterations);

    std::vector<double> alternating;
    std::vector<double> loadSamples;
    std::vector<double> releaseSamples;
    alternating.reserve(iterations);

    // A second texture, so the alternating leg exercises a genuinely different
    // cached image rather than the same one twice.
    SAFEARRAYBOUND secondBound{static_cast<ULONG>(contents.size()), 0};
    SAFEARRAY* secondBytes = SafeArrayCreate(VT_UI1, 1, &secondBound);
    if (!secondBytes) {
        throw std::bad_alloc();
    }
    ArrayGuard secondGuard{secondBytes};
    void* secondRaw = nullptr;
    bb::check(SafeArrayAccessData(secondBytes, &secondRaw), "SafeArrayAccessData");
    std::memcpy(secondRaw, contents.data(), contents.size());
    SafeArrayUnaccessData(secondBytes);

    long handle = 0;
    long other = 0;
    double loadMs = 0.0;
    try {
        bb::Value fill = bb::get(shape.obj(), L"Fill");

        // Warm both paths once; the first call of each pays one-off setup.
        bb::call(fill.obj(), L"UserPicture", {bb::Value(imagePath.c_str())});
        const long warmHandle = nativeTextureLoad(bytes);
        nativeTextureApply(fill.obj(), warmHandle);
        nativeTextureRelease(warmHandle);

        for (long index = 0; index < iterations; ++index) {
            const long long start = Now();
            bb::call(fill.obj(), L"UserPicture", {bb::Value(imagePath.c_str())});
            userPicture.push_back((Now() - start) * tick * 1000.0);
        }

        const long long loadStart = Now();
        handle = nativeTextureLoad(bytes);
        loadMs = (Now() - loadStart) * tick * 1000.0;

        for (long index = 0; index < iterations; ++index) {
            const long long start = Now();
            nativeTextureApply(fill.obj(), handle);
            applyTexture.push_back((Now() - start) * tick * 1000.0);
        }

        // Alternating two textures on one Shape: every apply genuinely changes
        // the fill, which is the pessimistic case for Office's own bookkeeping.
        other = nativeTextureLoad(secondBytes);
        for (long index = 0; index < iterations; ++index) {
            const long chosen = (index % 2 == 0) ? handle : other;
            const long long start = Now();
            nativeTextureApply(fill.obj(), chosen);
            alternating.push_back((Now() - start) * tick * 1000.0);
        }

        // Load and release cost, sampled over their own runs rather than once.
        constexpr long kLifecycleSamples = 50;
        for (long index = 0; index < kLifecycleSamples; ++index) {
            const long long loadedStart = Now();
            const long sample = nativeTextureLoad(bytes);
            loadSamples.push_back((Now() - loadedStart) * tick * 1000.0);
            const long long releaseStart = Now();
            nativeTextureRelease(sample);
            releaseSamples.push_back((Now() - releaseStart) * tick * 1000.0);
        }
    } catch (...) {
        if (other) {
            nativeTextureRelease(other);
        }
        if (handle) {
            nativeTextureRelease(handle);
        }
        bb::call(shape.obj(), L"Delete");
        throw;
    }

    const std::wstring cachedReport = nativeTextureReport(handle);
    nativeTextureRelease(other);
    nativeTextureRelease(handle);
    bb::call(shape.obj(), L"Delete");

    const Timing picture = Summarise(userPicture);
    const Timing apply = Summarise(applyTexture);
    const Timing alternate = Summarise(alternating);
    const Timing load = Summarise(loadSamples);
    const Timing release = Summarise(releaseSamples);

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(4);
    out << L"iterations=" << iterations << L';';
    Append(out, L"userPicture", picture);
    Append(out, L"applyTexture", apply);
    Append(out, L"applyAlternating", alternate);
    Append(out, L"loadTexture", load);
    Append(out, L"releaseTexture", release);
    out << L"firstLoadMs=" << loadMs << L';';
    if (apply.meanMs > 0.0) {
        out << L"meanSpeedup=" << (picture.meanMs / apply.meanMs) << L';';
    }
    if (apply.medianMs > 0.0) {
        out << L"medianSpeedup=" << (picture.medianMs / apply.medianMs) << L';';
    }
    if (alternate.meanMs > 0.0) {
        out << L"alternatingMeanSpeedup=" << (picture.meanMs / alternate.meanMs) << L';';
    }
    out << L"loadAmortisedOverCalls="
        << (apply.meanMs < picture.meanMs
                ? load.meanMs / (picture.meanMs - apply.meanMs)
                : 0.0)
        << L';';
    out << cachedReport;
    return out.str();
}

/**
 * Compares loading an encoded image against loading raw pixels, at one size.
 *
 * The encoded leg reads a real PNG from @p imagePath; the raw leg builds a BGRA
 * buffer of @p width x @p height. Both go through the public C ABI and release
 * each texture immediately, so what is timed is creation and teardown rather
 * than any steady-state cache.
 *
 * The two legs are not the same picture, and they are not meant to be: the
 * question is what a caller pays to get pixels it already holds into Office,
 * versus what it pays to hand over an encoded file of comparable size.
 */
std::wstring benchmarkPixelLoad(const std::wstring& imagePath, long width, long height,
                                long iterations) {
    if (width <= 0 || height <= 0 || iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Size and iterations must be positive");
    }
    if (BB_Init() != BB_OK) {
        char message[512]{};
        BB_GetLastError(message, sizeof(message));
        throw bb::Error(E_NOTIMPL, std::string("BB_Init failed: ") + message);
    }

    std::ifstream file(std::filesystem::path(imagePath), std::ios::binary);
    if (!file) {
        throw bb::Error(E_INVALIDARG, "Cannot read the benchmark image");
    }
    const std::vector<unsigned char> encoded((std::istreambuf_iterator<char>(file)),
                                             std::istreambuf_iterator<char>());

    // A gradient rather than a flat colour, so nothing can collapse the work.
    const long stride = width * 4;
    std::vector<unsigned char> pixels(static_cast<std::size_t>(stride) * height);
    for (long y = 0; y < height; ++y) {
        for (long x = 0; x < width; ++x) {
            unsigned char* pixel = pixels.data() + y * stride + x * 4;
            pixel[0] = static_cast<unsigned char>(x);
            pixel[1] = static_cast<unsigned char>(y);
            pixel[2] = static_cast<unsigned char>(x ^ y);
            pixel[3] = 0xFF;
        }
    }

    const double tick = SecondsPerTick();
    std::vector<double> encodedSamples;
    std::vector<double> pixelSamples;
    encodedSamples.reserve(iterations);
    pixelSamples.reserve(iterations);

    BB_Handle warm = 0;
    if (BB_LoadTexture(encoded.data(), static_cast<uint32_t>(encoded.size()), &warm) == BB_OK) {
        BB_ReleaseTexture(warm);
    }
    if (BB_LoadTexturePixels(pixels.data(), static_cast<uint32_t>(width),
                             static_cast<uint32_t>(height), static_cast<int32_t>(stride),
                             &warm) == BB_OK) {
        BB_ReleaseTexture(warm);
    }

    for (long index = 0; index < iterations; ++index) {
        BB_Handle handle = 0;
        const long long start = Now();
        const BB_Result status =
            BB_LoadTexture(encoded.data(), static_cast<uint32_t>(encoded.size()), &handle);
        encodedSamples.push_back((Now() - start) * tick * 1000.0);
        if (status != BB_OK) {
            throw bb::Error(E_FAIL, "BB_LoadTexture failed during the benchmark");
        }
        BB_ReleaseTexture(handle);
    }

    for (long index = 0; index < iterations; ++index) {
        BB_Handle handle = 0;
        const long long start = Now();
        const BB_Result status = BB_LoadTexturePixels(
            pixels.data(), static_cast<uint32_t>(width), static_cast<uint32_t>(height),
            static_cast<int32_t>(stride), &handle);
        pixelSamples.push_back((Now() - start) * tick * 1000.0);
        if (status != BB_OK) {
            char message[512]{};
            BB_GetLastError(message, sizeof(message));
            throw bb::Error(E_FAIL, std::string("BB_LoadTexturePixels failed: ") + message);
        }
        BB_ReleaseTexture(handle);
    }

    const Timing encodedTiming = Summarise(encodedSamples);
    const Timing pixelTiming = Summarise(pixelSamples);

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(4);
    out << L"size=" << width << L"x" << height << L";iterations=" << iterations
        << L";encodedBytes=" << encoded.size() << L";pixelBytes=" << pixels.size() << L';';
    Append(out, L"encodedLoad", encodedTiming);
    Append(out, L"pixelLoad", pixelTiming);
    if (pixelTiming.meanMs > 0.0) {
        out << L"pixelSpeedup=" << (encodedTiming.meanMs / pixelTiming.meanMs) << L';';
    }
    return out.str();
}
