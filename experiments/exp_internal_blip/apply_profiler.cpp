/**
 * @file apply_profiler.cpp
 * Where the ~0.186 ms of a cached ApplyTexture actually goes.
 *
 * The earlier stage profiler measured the *research* path - ApplyCachedImage
 * driven directly. This one walks the exact sequence `BB_ApplyTexture` walks,
 * stage by stage, so the number it produces is the number a caller pays:
 *
 *     BB_ApplyTexture
 *       RequireFillableShape
 *         VirtualQuery
 *         QueryInterface(IID_IDispatch)
 *         ClassifyShapeForNativePictureFill   (Shape.Type, Shape.Connector)
 *       get(shape, "Fill")                     (Automation property fetch)
 *       TextureStore lookup                    (handle -> TextureRef)
 *       ResolveFillTarget                      (modules, wrapper, receiver)
 *       ResolveApplyFunctions                  (signature-verified entry points)
 *       ApplyCachedImage
 *         record, subrecord, install, transfer, holder, transaction,
 *         apply, destructors
 *
 * Every stage is timed with QueryPerformanceCounter and accumulated into a
 * vector; nothing is printed inside the loop, because a single formatted write
 * costs more than most of the stages being measured.
 *
 * Reported per stage: mean, median, p95, p99, min, max. The median matters more
 * than the mean here - one scheduling stall in a few thousand iterations moves
 * the mean of a two-microsecond stage by more than the stage costs.
 *
 * Research only. STA, PowerPoint host, leaves the document as found.
 */

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../../src/backend/windows_office/shape_policy.hpp"
#include "../experiment_api.hpp"

#include <algorithm>
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cmath>
#include <cstring>
#include <cwchar>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

namespace {

constexpr long kDefaultIterations = 2000;

double SecondsPerTick() {
    LARGE_INTEGER frequency{};
    QueryPerformanceFrequency(&frequency);
    return 1.0 / static_cast<double>(frequency.QuadPart);
}

inline long long Now() {
    LARGE_INTEGER counter{};
    QueryPerformanceCounter(&counter);
    return counter.QuadPart;
}

/// Nearest-rank percentile on an already sorted sample.
double Percentile(const std::vector<double>& sorted, double fraction) {
    if (sorted.empty()) {
        return 0.0;
    }
    const auto rank = static_cast<std::size_t>(fraction * static_cast<double>(sorted.size()));
    return sorted[std::min(rank, sorted.size() - 1)];
}

/**
 * One named stage's samples.
 *
 * Samples are reserved up front so a reallocation cannot land inside a timed
 * region and show up as that stage being slow.
 */
class Stage {
  public:
    void Reserve(std::size_t count) {
        samples_.reserve(count);
    }

    inline void Add(double milliseconds) {
        samples_.push_back(milliseconds);
    }

    std::size_t Count() const {
        return samples_.size();
    }

    void Write(std::wostringstream& out, const wchar_t* leg, const wchar_t* name) {
        if (samples_.empty()) {
            return;
        }
        std::vector<double> sorted = samples_;
        std::sort(sorted.begin(), sorted.end());
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        out << leg << L'.' << name << L".mean=" << (total / samples_.size()) << L';' << leg << L'.'
            << name << L".median=" << Percentile(sorted, 0.50) << L';' << leg << L'.' << name
            << L".p95=" << Percentile(sorted, 0.95) << L';' << leg << L'.' << name
            << L".p99=" << Percentile(sorted, 0.99) << L';' << leg << L'.' << name
            << L".min=" << sorted.front() << L';' << leg << L'.' << name
            << L".max=" << sorted.back() << L';';
    }

    double Mean() const {
        if (samples_.empty()) {
            return 0.0;
        }
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        return total / samples_.size();
    }

  private:
    std::vector<double> samples_;
};

/// The stages of the apply, in call order so the report reads like the path.
struct ApplyStages {
    Stage virtualQuery;
    Stage queryInterface;
    Stage classify;
    Stage fillFetch;
    Stage handleLookup;
    Stage resolveTarget;
    Stage resolveFunctions;
    Stage record;
    Stage subrecord;
    Stage install;
    Stage transfer;
    Stage holder;
    Stage transaction;
    Stage nativeApply;
    Stage destructors;
    Stage total;

    void Reserve(std::size_t count) {
        for (Stage* stage : All()) {
            stage->Reserve(count);
        }
    }

    std::vector<Stage*> All() {
        return {&virtualQuery, &queryInterface, &classify,    &fillFetch,
                &handleLookup, &resolveTarget,  &resolveFunctions, &record,
                &subrecord,    &install,        &transfer,    &holder,
                &transaction,  &nativeApply,    &destructors, &total};
    }

    std::vector<const wchar_t*> Names() const {
        return {L"virtualQuery", L"queryInterface", L"classify",    L"fillFetch",
                L"handleLookup", L"resolveTarget",  L"resolveFunctions", L"record",
                L"subrecord",    L"install",        L"transfer",    L"holder",
                L"transaction",  L"nativeApply",    L"destructors", L"total"};
    }

    void Write(std::wostringstream& out, const wchar_t* leg) {
        std::vector<Stage*> stages = All();
        std::vector<const wchar_t*> names = Names();
        double sum = 0.0;
        for (std::size_t i = 0; i < stages.size(); ++i) {
            stages[i]->Write(out, leg, names[i]);
            if (names[i] != std::wstring(L"total")) {
                sum += stages[i]->Mean();
            }
        }
        // The stage means should add up to the measured total; a gap is time the
        // instrumentation did not attribute and must be visible, not hidden.
        out << leg << L".stageSum=" << sum << L';' << leg
            << L".unattributed=" << (total.Mean() - sum) << L';';
    }
};

/**
 * Routes ApplyCachedImage's stage hook into the timing vectors.
 *
 * The hook is called with a stage name after each internal step; the elapsed
 * time since the previous checkpoint is that step's cost. `enter` sets the
 * baseline and is not itself a stage.
 */
class InnerTimer {
  public:
    InnerTimer(ApplyStages& stages, double tick) : stages_(stages), tick_(tick) {}

    bb::oart::StageSampler Sampler() {
        return [this](const wchar_t* stage) {
            const long long now = Now();
            if (std::wcscmp(stage, L"enter") == 0) {
                previous_ = now;
                return;
            }
            const double ms = static_cast<double>(now - previous_) * tick_ * 1000.0;
            previous_ = now;
            if (std::wcscmp(stage, L"record") == 0) {
                stages_.record.Add(ms);
            } else if (std::wcscmp(stage, L"subrecord") == 0) {
                stages_.subrecord.Add(ms);
            } else if (std::wcscmp(stage, L"install") == 0) {
                stages_.install.Add(ms);
            } else if (std::wcscmp(stage, L"transfer") == 0) {
                stages_.transfer.Add(ms);
            } else if (std::wcscmp(stage, L"holder") == 0) {
                stages_.holder.Add(ms);
            } else if (std::wcscmp(stage, L"transaction") == 0) {
                stages_.transaction.Add(ms);
            } else if (std::wcscmp(stage, L"apply") == 0) {
                stages_.nativeApply.Add(ms);
            } else if (std::wcscmp(stage, L"released") == 0) {
                stages_.destructors.Add(ms);
            }
        };
    }

  private:
    ApplyStages& stages_;
    double tick_ = 0.0;
    long long previous_ = 0;
};

/// Releases the QueryInterface reference the way the production path does.
struct DispatchGuard {
    IDispatch* value = nullptr;

    ~DispatchGuard() {
        if (value) {
            value->Release();
        }
    }
};

} // namespace

/**
 * Profiles the production apply path on one Shape, @p iterations times.
 *
 * @p shape is a live PowerPoint Shape and @p handle a texture already loaded, so
 * nothing here measures decoding. The document is left as it was found apart
 * from the Shape's fill, which is what the apply is for.
 */
std::wstring profileApplyStages(IDispatch* shape, long handle, long iterations) {
    if (iterations <= 0) {
        iterations = kDefaultIterations;
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "The apply profiler requires the PowerPoint host");
    }
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }

    const double tick = SecondsPerTick();
    ApplyStages stages;
    stages.Reserve(static_cast<std::size_t>(iterations));

    // Warm every cache the production path keeps: module validation, wrapper
    // decoding, entry-point signatures. The first call pays all of it, and
    // leaving that in the samples would misattribute one-off cost to a stage.
    {
        IDispatch* warm = nullptr;
        shape->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&warm));
        DispatchGuard guard{warm};
        bb::office::RequireNativePictureFillTarget(warm);
        bb::Value fill = bb::get(warm, L"Fill");
        nativeTextureApply(fill.obj(), handle);
    }

    InnerTimer timer(stages, tick);
    const bb::oart::StageSampler sampler = timer.Sampler();

    for (long index = 0; index < iterations; ++index) {
        const long long begin = Now();

        // --- RequireFillableShape, split into its parts ----------------------
        long long mark = begin;
        MEMORY_BASIC_INFORMATION information{};
        VirtualQuery(shape, &information, sizeof(information));
        long long now = Now();
        stages.virtualQuery.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        IDispatch* dispatch = nullptr;
        shape->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&dispatch));
        DispatchGuard guard{dispatch};
        now = Now();
        stages.queryInterface.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        bb::office::RequireNativePictureFillTarget(dispatch);
        now = Now();
        stages.classify.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        // --- the Automation fetch of Shape.Fill -------------------------------
        bb::Value fill = bb::get(dispatch, L"Fill");
        now = Now();
        stages.fillFetch.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        // --- handle to image --------------------------------------------------
        const bb::office::TextureRef texture = bb::office::LookupTexture(handle);
        now = Now();
        stages.handleLookup.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        // --- the guard chain ---------------------------------------------------
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
        now = Now();
        stages.resolveTarget.Add(static_cast<double>(now - mark) * tick * 1000.0);
        mark = now;

        const bb::oart::ApplyFunctions functions =
            bb::oart::ResolveApplyFunctions(target.oartBase);
        now = Now();
        stages.resolveFunctions.Add(static_cast<double>(now - mark) * tick * 1000.0);

        // --- the apply itself, with its internal stages -----------------------
        bb::oart::ApplyCachedImage(functions, target, bb::office::CachedImageOf(texture), sampler);

        stages.total.Add(static_cast<double>(Now() - begin) * tick * 1000.0);
    }

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    out << L"iterations=" << iterations << L';';
    stages.Write(out, L"apply");
    return out.str();
}
