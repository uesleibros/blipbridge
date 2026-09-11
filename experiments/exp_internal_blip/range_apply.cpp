/**
 * @file range_apply.cpp
 * Benchmarks the three ways to fill N Shapes with one texture.
 *
 * All three go through the **public C ABI**, in this process, so what is
 * compared is what a caller actually gets:
 *
 *   - N x `BB_ApplyTexture`   - one apply per Shape
 *   - `BB_ApplyTextureBatch`  - one ABI crossing, still N applies
 *   - `BB_ApplyTextureRange`  - one apply through Office's own range receiver
 *
 * Both comparison legs have to run in this process. Driven from PowerShell they
 * read about 5.9 ms per Shape, which is a cross-process Automation round trip
 * rather than Office, and would make the range look twenty-five times better
 * than it is.
 *
 * The implementation under test lives in
 * `src/backend/windows_office/range_texture.cpp`, which also explains why the
 * API takes a ShapeRange rather than an array of Shapes. Nothing here
 * reimplements it.
 */

#include "../experiment_api.hpp"

#include <algorithm>
#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <sstream>
#include <string>
#include <vector>

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

[[noreturn]] void Fail(const char* what) {
    char message[512]{};
    BB_GetLastError(message, sizeof(message));
    throw bb::Error(E_FAIL, std::string(what) + ": " + message);
}

/// One timed leg. Samples reserved up front, summarised after the loop.
class Leg {
  public:
    explicit Leg(long capacity) {
        samples_.reserve(static_cast<std::size_t>(capacity));
    }

    void Add(double ms) {
        samples_.push_back(ms);
    }

    double MeanMs() const {
        if (samples_.empty()) {
            return 0.0;
        }
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        return total / static_cast<double>(samples_.size());
    }

    void Write(std::wostringstream& out, const wchar_t* name) {
        if (samples_.empty()) {
            return;
        }
        std::sort(samples_.begin(), samples_.end());
        out << name << L"MeanMs=" << MeanMs() << L';' << name << L"MedianMs="
            << samples_[samples_.size() / 2] << L';' << name << L"MinMs=" << samples_.front()
            << L';';
    }

  private:
    std::vector<double> samples_;
};

} // namespace

/**
 * Times the three routes over @p range, @p iterations rounds each.
 *
 * The members are read once for the per-Shape legs; the range apply is handed
 * the ShapeRange itself, which is the whole point of it.
 */
std::wstring applyTextureToRange(IDispatch* range, long handle, long iterations) {
    if (!range) {
        throw bb::Error(E_POINTER, "Missing ShapeRange");
    }
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }
    if (BB_Init() != BB_OK) {
        Fail("BB_Init failed");
    }
    const auto texture = static_cast<BB_Handle>(handle);

    // The members, held for the per-Shape legs. The Values keep them alive for
    // the duration of the call and nothing is retained afterwards.
    const long count = bb::get(range, L"Count").integer();
    std::vector<bb::Value> members;
    std::vector<void*> pointers;
    members.reserve(static_cast<std::size_t>(count));
    pointers.reserve(static_cast<std::size_t>(count));
    for (long index = 1; index <= count; ++index) {
        bb::Value member = bb::call(range, L"Item", {bb::Value(index)});
        pointers.push_back(member.obj());
        members.push_back(std::move(member));
    }
    std::vector<BB_Handle> handles(static_cast<std::size_t>(count), texture);

    const double tick = SecondsPerTick();
    Leg loop(iterations);
    Leg batch(iterations);
    Leg ranged(iterations);

    // Warm all three, so none pays first-call setup inside its samples.
    std::uint32_t applied = 0;
    for (void* shape : pointers) {
        if (BB_ApplyTexture(shape, texture) != BB_OK) {
            Fail("BB_ApplyTexture failed while warming");
        }
    }
    if (BB_ApplyTextureBatch(
            pointers.data(), handles.data(), static_cast<uint32_t>(count), &applied) != BB_OK) {
        Fail("BB_ApplyTextureBatch failed while warming");
    }
    if (BB_ApplyTextureRange(range, texture, &applied) != BB_OK) {
        Fail("BB_ApplyTextureRange failed while warming");
    }
    if (applied != static_cast<std::uint32_t>(count)) {
        throw bb::Error(E_FAIL, "The range apply reported fewer Shapes than the range holds");
    }

    for (long round = 0; round < iterations; ++round) {
        long long start = Now();
        for (void* shape : pointers) {
            if (BB_ApplyTexture(shape, texture) != BB_OK) {
                Fail("BB_ApplyTexture failed during the benchmark");
            }
        }
        loop.Add(static_cast<double>(Now() - start) * tick * 1000.0);

        start = Now();
        if (BB_ApplyTextureBatch(
                pointers.data(), handles.data(), static_cast<uint32_t>(count), &applied) !=
            BB_OK) {
            Fail("BB_ApplyTextureBatch failed during the benchmark");
        }
        batch.Add(static_cast<double>(Now() - start) * tick * 1000.0);

        start = Now();
        if (BB_ApplyTextureRange(range, texture, &applied) != BB_OK) {
            Fail("BB_ApplyTextureRange failed during the benchmark");
        }
        ranged.Add(static_cast<double>(Now() - start) * tick * 1000.0);
    }

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    const auto shapes = static_cast<double>(count);
    out << L"iterations=" << iterations << L";shapes=" << count << L';';
    loop.Write(out, L"loop");
    batch.Write(out, L"batch");
    ranged.Write(out, L"range");
    const double rangeMean = ranged.MeanMs();
    out << L"perShapeLoopMs=" << (loop.MeanMs() / shapes) << L';' << L"perShapeBatchMs="
        << (batch.MeanMs() / shapes) << L';' << L"perShapeRangeMs=" << (rangeMean / shapes) << L';';
    if (rangeMean > 0.0) {
        out << L"speedupVsLoop=" << (loop.MeanMs() / rangeMean) << L';' << L"speedupVsBatch="
            << (batch.MeanMs() / rangeMean) << L';';
    }
    return out.str();
}
