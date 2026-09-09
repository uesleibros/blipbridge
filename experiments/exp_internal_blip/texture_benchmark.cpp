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

/// Mean and median of per-call milliseconds, which is what the report quotes.
struct Timing {
    double meanMs = 0.0;
    double medianMs = 0.0;
    double totalMs = 0.0;
    int calls = 0;
};

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
    timing.medianMs = samples[samples.size() / 2];
    return timing;
}

void Append(std::wostringstream& out, const wchar_t* label, const Timing& timing) {
    out << label << L"Calls=" << timing.calls << L';' << label << L"MeanMs="
        << timing.meanMs << L';' << label << L"MedianMs=" << timing.medianMs << L';'
        << label << L"TotalMs=" << timing.totalMs << L';';
}

} // namespace

/**
 * Benchmarks @p iterations fills three ways on the Shapes of @p slide.
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
    struct ArrayGuard {
        SAFEARRAY* value;
        ~ArrayGuard() { SafeArrayDestroy(value); }
    } arrayGuard{bytes};
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

    long handle = 0;
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
    } catch (...) {
        if (handle) {
            nativeTextureRelease(handle);
        }
        bb::call(shape.obj(), L"Delete");
        throw;
    }

    const std::wstring cachedReport = nativeTextureReport(handle);
    nativeTextureRelease(handle);
    bb::call(shape.obj(), L"Delete");

    const Timing picture = Summarise(userPicture);
    const Timing apply = Summarise(applyTexture);

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(4);
    out << L"iterations=" << iterations << L';';
    Append(out, L"userPicture", picture);
    out << L"loadTextureMs=" << loadMs << L';';
    Append(out, L"applyTexture", apply);
    if (apply.meanMs > 0.0) {
        out << L"meanSpeedup=" << (picture.meanMs / apply.meanMs) << L';';
    }
    if (apply.medianMs > 0.0) {
        out << L"medianSpeedup=" << (picture.medianMs / apply.medianMs) << L';';
    }
    out << L"loadAmortisedOverCalls="
        << (apply.meanMs < picture.meanMs
                ? loadMs / (picture.meanMs - apply.meanMs)
                : 0.0)
        << L';';
    out << cachedReport;
    return out.str();
}
