/**
 * @file transaction_split.cpp
 * Which half of the private apply is the expensive one.
 *
 * The receiver's entry point does two things with a finished transaction: it
 * performs the change, and then it records that change against the document.
 * Everything about whether a multi-Shape batch could ever exist turns on which
 * of those costs the 0.15 ms, because a batch can only ever share the second
 * one - the change is per Shape by construction, since the transaction carries
 * no target and the receiver *is* the target.
 *
 * So this runs the same apply both ways on the same Shape: through the
 * receiver's own entry point, and through the two steps taken apart. It reports
 * the halves, and it reports the two totals side by side - if the split route
 * were not a faithful reproduction, its total would not match.
 *
 * Research only. Nothing here is reachable from the C ABI.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../../src/backend/windows_office/shape_policy.hpp"

#include <algorithm>
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

/// Samples in milliseconds, summarised after the loop and never inside it.
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

    double MedianMs() {
        if (samples_.empty()) {
            return 0.0;
        }
        std::sort(samples_.begin(), samples_.end());
        return samples_[samples_.size() / 2];
    }

    void Write(std::wostringstream& out, const wchar_t* name) {
        out << name << L"MeanMs=" << MeanMs() << L';' << name << L"MedianMs=" << MedianMs() << L';';
    }

  private:
    std::vector<double> samples_;
};

} // namespace

/**
 * Applies @p handle straight to a FillFormat, @p iterations times, and reports it.
 *
 * The point is what can be passed in: `ShapeRange.Fill` is the same PPCORE
 * delegating wrapper as `Shape.Fill`, so this can ask whether one private apply
 * against a range of Shapes fills all of them - which is the only shape a true
 * multi-Shape transaction could take, given that a transaction carries no target
 * and the receiver *is* the target.
 *
 * Research only, and it deliberately skips the semantic gate, because the gate
 * asks about a Shape and this is handed a fill. That makes it the caller's job
 * to pass a range of classes the gate would accept - which is fine for a
 * measurement and is exactly why this is not an API.
 */
std::wstring applyCachedImageToFill(IDispatch* fill, long handle, long iterations) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing FillFormat");
    }
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }
    const bb::office::TextureRef texture = bb::office::LookupTexture(handle);
    void* cached = bb::office::CachedImageOf(texture);
    if (!cached) {
        throw bb::Error(E_FAIL, "No image behind that handle");
    }

    const double tick = SecondsPerTick();
    Leg applies(iterations);

    const bb::oart::FillTarget warm = bb::oart::ResolveFillTarget(fill);
    const bb::oart::ApplyFunctions functions = bb::oart::ResolveApplyFunctions(warm.oartBase);
    bb::oart::ApplyCachedImage(functions, warm, cached);

    for (long round = 0; round < iterations; ++round) {
        // Re-resolved every round, exactly as the production path does: a fill
        // whose Shape went away must not be reached through a kept pointer.
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill);
        const long long start = Now();
        bb::oart::ApplyCachedImage(functions, target, cached);
        applies.Add(static_cast<double>(Now() - start) * tick * 1000.0);
    }

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    out << L"iterations=" << iterations << L';';
    applies.Write(out, L"apply");
    return out.str();
}

/**
 * Applies @p handle to @p shape once, performing the change without recording it.
 *
 * Research only, and deliberately awkward to reach: this leaves the document
 * changed with nothing told about it. What that costs a document is the whole
 * question tools/test_change_only.ps1 exists to answer.
 */
std::wstring applyChangeOnly(IDispatch* shape, long handle) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    bb::office::RequireNativePictureFillTarget(shape);
    const bb::office::TextureRef texture = bb::office::LookupTexture(handle);
    void* cached = bb::office::CachedImageOf(texture);
    if (!cached) {
        throw bb::Error(E_FAIL, "No image behind that handle");
    }

    bb::Value fill = bb::get(shape, L"Fill");
    const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
    const bb::oart::ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);
    bb::oart::ApplyCachedImage(functions, target, cached, {}, bb::oart::ApplyRoute::ChangeOnly);
    return L"applied=1;recorded=0;";
}

/**
 * Applies @p handle to @p shape @p iterations times each way and reports both.
 *
 * The Shape must already be one the semantic gate accepts; it is checked here
 * the same way the production path checks it, because this reaches the private
 * backend directly and the gate is what stands in front of it.
 */
std::wstring splitTransactionApply(IDispatch* shape, long handle, long iterations) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }
    bb::office::RequireNativePictureFillTarget(shape);

    const bb::office::TextureRef texture = bb::office::LookupTexture(handle);
    void* cached = bb::office::CachedImageOf(texture);
    if (!cached) {
        throw bb::Error(E_FAIL, "No image behind that handle");
    }

    const double tick = SecondsPerTick();
    Leg combined(iterations);
    Leg split(iterations);
    Leg change(iterations);
    Leg record(iterations);
    Leg changeOnly(iterations);

    // Timed inside the sampler, which fires after each step. The transaction
    // stage is the baseline the two halves are measured from.
    long long stageStart = 0;
    double changeMs = 0.0;
    double recordMs = 0.0;
    const bb::oart::StageSampler sampler = [&](const wchar_t* stage) {
        const long long now = Now();
        const double elapsed = static_cast<double>(now - stageStart) * tick * 1000.0;
        if (std::wstring(stage) == L"change") {
            changeMs = elapsed;
        } else if (std::wstring(stage) == L"record") {
            recordMs = elapsed;
        }
        stageStart = now;
    };

    // Warm both routes so neither pays for first-call work in its samples.
    {
        bb::Value fill = bb::get(shape, L"Fill");
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
        const bb::oart::ApplyFunctions functions =
            bb::oart::ResolveApplyFunctions(target.oartBase);
        bb::oart::ApplyCachedImage(functions, target, cached, {}, bb::oart::ApplyRoute::Combined);
        bb::oart::ApplyCachedImage(functions, target, cached, {}, bb::oart::ApplyRoute::Split);
        bb::oart::ApplyCachedImage(
            functions, target, cached, {}, bb::oart::ApplyRoute::ChangeOnly);
    }

    for (long round = 0; round < iterations; ++round) {
        bb::Value fill = bb::get(shape, L"Fill");
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
        const bb::oart::ApplyFunctions functions =
            bb::oart::ResolveApplyFunctions(target.oartBase);

        long long start = Now();
        bb::oart::ApplyCachedImage(functions, target, cached, {}, bb::oart::ApplyRoute::Combined);
        combined.Add(static_cast<double>(Now() - start) * tick * 1000.0);

        changeMs = 0.0;
        recordMs = 0.0;
        start = Now();
        stageStart = start;
        bb::oart::ApplyCachedImage(
            functions, target, cached, sampler, bb::oart::ApplyRoute::Split);
        split.Add(static_cast<double>(Now() - start) * tick * 1000.0);
        change.Add(changeMs);
        record.Add(recordMs);

        // The floor: the same change with nothing told about it. What a picture
        // fill would cost if the recording were free.
        start = Now();
        bb::oart::ApplyCachedImage(
            functions, target, cached, {}, bb::oart::ApplyRoute::ChangeOnly);
        changeOnly.Add(static_cast<double>(Now() - start) * tick * 1000.0);
    }

    // The document has to agree that both routes did the same thing.
    bb::Value fill = bb::get(shape, L"Fill");
    const long fillType = bb::get(fill.obj(), L"Type").integer();

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    out << L"iterations=" << iterations << L";fillType=" << fillType << L';';
    combined.Write(out, L"combined");
    split.Write(out, L"split");
    change.Write(out, L"change");
    record.Write(out, L"record");
    changeOnly.Write(out, L"changeOnly");
    const double splitMean = split.MeanMs();
    if (splitMean > 0.0) {
        out << L"changeShare=" << (100.0 * change.MeanMs() / splitMean) << L';'
            << L"recordShare=" << (100.0 * record.MeanMs() / splitMean) << L';';
    }
    out << L"routeDifferenceMs=" << (splitMean - combined.MeanMs()) << L';';
    return out.str();
}
