/**
 * @file skip_benchmark.cpp
 * What BB_ApplyTextureIfChanged actually costs, in the six shapes of use that
 * decide whether it is worth having.
 *
 * The profiler put 76% of an apply in one private OART call - the document edit
 * itself. Nothing inside that call can be made cheaper from outside Office, so
 * the only way past it is to not make it. This measures whether "not making it"
 * is genuinely cheap once the cost of *proving* it is safe to skip is included,
 * because that proof - a Shape key and a re-read of Fill.Type - is not free and
 * is what the feature lives or dies by.
 *
 * Everything runs through the real C ABI in process, on real Shapes on the
 * caller's slide, and every leg asserts the skip decision it expected rather
 * than merely timing whatever happened. A benchmark that silently stopped
 * skipping would otherwise look like a triumph.
 */

#include "../experiment_api.hpp"

#include <algorithm>
#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>
#include <cmath>
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

/// The C ABI's own account of why it failed, so a benchmark failure is legible.
[[noreturn]] void Fail(const char* what) {
    char message[512]{};
    BB_GetLastError(message, sizeof(message));
    throw bb::Error(E_FAIL, std::string(what) + ": " + message);
}

/**
 * One timed leg.
 *
 * Samples are reserved up front and never printed inside the loop: a single
 * formatting call between samples would cost more than the skip being measured.
 */
class Leg {
  public:
    Leg(std::wstring name, long capacity) : name_(std::move(name)) {
        samples_.reserve(static_cast<std::size_t>(capacity));
    }

    void Add(double ms) {
        samples_.push_back(ms);
    }

    void Write(std::wostringstream& out) {
        if (samples_.empty()) {
            return;
        }
        std::sort(samples_.begin(), samples_.end());
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        const std::size_t count = samples_.size();
        const auto at = [&](double quantile) {
            auto index = static_cast<std::size_t>(quantile * static_cast<double>(count));
            return samples_[index >= count ? count - 1 : index];
        };
        out << name_ << L"MeanMs=" << (total / static_cast<double>(count)) << L';' << name_
            << L"MedianMs=" << at(0.50) << L';' << name_ << L"P95Ms=" << at(0.95) << L';' << name_
            << L"P99Ms=" << at(0.99) << L';' << name_ << L"MinMs=" << samples_.front() << L';'
            << name_ << L"MaxMs=" << samples_.back() << L';' << name_ << L"Samples=" << count
            << L';';
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

  private:
    std::wstring name_;
    std::vector<double> samples_;
};

/// A 1x1 PNG, so the benchmark carries its own image and decodes it once.
const unsigned char kRedPixelPng[] = {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
    0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
    0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
    0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D,
    0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82};

/// A second 1x1 PNG with different content, so alternating really alternates.
const unsigned char kBluePixelPng[] = {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
    0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
    0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
    0xD7, 0x63, 0x60, 0x60, 0xF8, 0x0F, 0x00, 0x01, 0x04, 0x01, 0x00, 0x8D, 0x9F, 0x1B,
    0xC6, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82};

BB_Handle LoadTexture(const unsigned char* bytes, std::size_t length) {
    BB_Handle handle = 0;
    if (BB_LoadTexture(bytes, static_cast<uint32_t>(length), &handle) != BB_OK) {
        Fail("BB_LoadTexture failed");
    }
    return handle;
}

/// Creates a small AutoShape and returns it, owning the reference.
bb::Value AddShape(IDispatch* shapes, long index) {
    const double left = 10.0 + static_cast<double>(index % 20) * 24.0;
    const double top = 10.0 + static_cast<double>(index / 20) * 24.0;
    return bb::call(
        shapes,
        L"AddShape",
        {bb::Value(1L), bb::Value(left), bb::Value(top), bb::Value(20.0), bb::Value(20.0)});
}

} // namespace

/**
 * Runs the six skip scenarios and reports what each one measured.
 *
 * @p shapeCount sizes the many-Shapes leg; @p iterations sizes every per-call
 * leg. The Shapes are created and deleted here, so the slide is left as found.
 */
std::wstring benchmarkApplySkip(IDispatch* slide, long shapeCount, long iterations) {
    if (shapeCount <= 0 || iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Shape count and iterations must be positive");
    }
    if (BB_Init() != BB_OK) {
        Fail("BB_Init failed");
    }

    bb::Value shapes = bb::get(slide, L"Shapes");
    std::vector<bb::Value> created;
    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);

    BB_Handle red = 0;
    BB_Handle blue = 0;
    try {
        red = LoadTexture(kRedPixelPng, sizeof(kRedPixelPng));
        blue = LoadTexture(kBluePixelPng, sizeof(kBluePixelPng));
        const double tick = SecondsPerTick();

        created.push_back(AddShape(shapes.obj(), 0));
        IDispatch* target = created.back().obj();

        // ---------------------------------------------------------------- 1 --
        // The baseline this whole phase is trying to beat: the shipping apply,
        // repeated, with nothing about the document changing.
        Leg baseline(L"baseline", iterations);
        if (BB_ApplyTexture(target, red) != BB_OK) {
            Fail("warm-up BB_ApplyTexture failed");
        }
        for (long round = 0; round < iterations; ++round) {
            const long long start = Now();
            const BB_Result result = BB_ApplyTexture(target, red);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTexture failed during the baseline");
            }
            baseline.Add(ms);
        }

        // ---------------------------------------------------------------- 2 --
        // IfChanged when it cannot skip. The Shape is invalidated first, outside
        // the timed region, so each sample is a genuine first apply: the cache
        // lookup misses and the full Office edit runs.
        Leg cold(L"ifChangedFirst", iterations);
        long coldSkips = 0;
        for (long round = 0; round < iterations; ++round) {
            if (BB_InvalidateShape(target) != BB_OK) {
                Fail("BB_InvalidateShape failed");
            }
            int32_t skipped = 0;
            const long long start = Now();
            const BB_Result result = BB_ApplyTextureIfChanged(target, red, &skipped);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTextureIfChanged failed on a first apply");
            }
            coldSkips += skipped;
            cold.Add(ms);
        }
        // An invalidated Shape must never be skipped: that is what invalidation
        // means, and a benchmark that let it slide would be measuring nothing.
        if (coldSkips != 0) {
            throw bb::Error(E_FAIL, "An invalidated Shape was skipped");
        }

        // ---------------------------------------------------------------- 3 --
        // The case the feature exists for: the same image, over and over.
        Leg hot(L"ifChangedRepeated", iterations);
        long hotApplies = 0;
        for (long round = 0; round < iterations; ++round) {
            int32_t skipped = 0;
            const long long start = Now();
            const BB_Result result = BB_ApplyTextureIfChanged(target, red, &skipped);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTextureIfChanged failed on a repeat");
            }
            hotApplies += skipped ? 0 : 1;
            hot.Add(ms);
        }
        if (hotApplies != 0) {
            throw bb::Error(E_FAIL, "A repeated identical apply was not skipped");
        }

        // ---------------------------------------------------------------- 4 --
        // The worst case: two images in turn, so the cache is consulted every
        // time and is never right. This is the overhead the feature charges a
        // caller who gains nothing from it, and it has to be small.
        Leg alternatingIfChanged(L"alternatingIfChanged", iterations);
        long alternatingSkips = 0;
        // Primed with the *other* image, or the first round would legitimately
        // skip: the Shape is still carrying what leg 3 left on it.
        if (BB_ApplyTexture(target, blue) != BB_OK) {
            Fail("BB_ApplyTexture failed while priming the alternating leg");
        }
        for (long round = 0; round < iterations; ++round) {
            const BB_Handle texture = (round % 2) ? blue : red;
            int32_t skipped = 0;
            const long long start = Now();
            const BB_Result result = BB_ApplyTextureIfChanged(target, texture, &skipped);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTextureIfChanged failed while alternating");
            }
            alternatingSkips += skipped;
            alternatingIfChanged.Add(ms);
        }
        if (alternatingSkips != 0) {
            throw bb::Error(E_FAIL, "A changed image was skipped");
        }

        Leg alternatingPlain(L"alternatingPlain", iterations);
        for (long round = 0; round < iterations; ++round) {
            const BB_Handle texture = (round % 2) ? blue : red;
            const long long start = Now();
            const BB_Result result = BB_ApplyTexture(target, texture);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTexture failed while alternating");
            }
            alternatingPlain.Add(ms);
        }

        // ---------------------------------------------------------------- 5 --
        // Deletion and recreation, which is the scenario that could make a
        // value-keyed cache dangerous: if PowerPoint reissues a deleted Shape's
        // Id, a new Shape inherits the old one's key. Both halves are recorded -
        // how often an Id was reused, and whether any reuse produced a skip.
        //
        // A fresh AutoShape has a solid fill, so the Fill.Type re-read refuses
        // the skip even on a reused Id. This leg is what turns that reasoning
        // into a measurement.
        const long lifecycleRounds = std::min<long>(iterations, 200);
        Leg lifecycle(L"deleteRecreate", lifecycleRounds);
        long reusedIds = 0;
        long wrongSkips = 0;
        long previousId = bb::get(target, L"Id").integer();
        for (long round = 0; round < lifecycleRounds; ++round) {
            bb::call(created.back().obj(), L"Delete");
            created.pop_back();

            created.push_back(AddShape(shapes.obj(), 0));
            target = created.back().obj();

            const long freshId = bb::get(target, L"Id").integer();
            if (freshId == previousId) {
                ++reusedIds;
            }
            previousId = freshId;

            int32_t skipped = 0;
            const long long start = Now();
            const BB_Result result = BB_ApplyTextureIfChanged(target, red, &skipped);
            const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
            if (result != BB_OK) {
                Fail("BB_ApplyTextureIfChanged failed on a recreated Shape");
            }
            wrongSkips += skipped;
            lifecycle.Add(ms);
        }

        // ---------------------------------------------------------------- 6 --
        // Many Shapes sharing one texture - the shape of a real slide redraw.
        // Timed as whole passes, because that is the unit a caller cares about.
        for (long index = 1; index < shapeCount; ++index) {
            created.push_back(AddShape(shapes.obj(), index));
        }
        std::vector<void*> pointers;
        pointers.reserve(created.size());
        for (bb::Value& shape : created) {
            pointers.push_back(shape.obj());
        }

        const long passes = std::max<long>(1, std::min<long>(iterations, 2000 / shapeCount + 1));
        for (void* shape : pointers) {
            if (BB_ApplyTexture(shape, red) != BB_OK) {
                Fail("BB_ApplyTexture failed while warming the many-Shape leg");
            }
            int32_t skipped = 0;
            if (BB_ApplyTextureIfChanged(shape, red, &skipped) != BB_OK) {
                Fail("BB_ApplyTextureIfChanged failed while warming the many-Shape leg");
            }
        }

        Leg manyPlain(L"manyShapesPlain", passes);
        for (long round = 0; round < passes; ++round) {
            const long long start = Now();
            for (void* shape : pointers) {
                if (BB_ApplyTexture(shape, red) != BB_OK) {
                    Fail("BB_ApplyTexture failed during the many-Shape pass");
                }
            }
            manyPlain.Add(static_cast<double>(Now() - start) * tick * 1000.0);
        }

        Leg manySkipped(L"manyShapesIfChanged", passes);
        long manyApplies = 0;
        for (long round = 0; round < passes; ++round) {
            const long long start = Now();
            for (void* shape : pointers) {
                int32_t skipped = 0;
                if (BB_ApplyTextureIfChanged(shape, red, &skipped) != BB_OK) {
                    Fail("BB_ApplyTextureIfChanged failed during the many-Shape pass");
                }
                manyApplies += skipped ? 0 : 1;
            }
            manySkipped.Add(static_cast<double>(Now() - start) * tick * 1000.0);
        }
        if (manyApplies != 0) {
            throw bb::Error(E_FAIL, "A Shape in the many-Shape pass was not skipped");
        }

        const std::size_t shapesInPass = pointers.size();
        out << L"iterations=" << iterations << L";shapeCount=" << shapeCount << L";passes="
            << passes << L';';
        baseline.Write(out);
        cold.Write(out);
        hot.Write(out);
        alternatingIfChanged.Write(out);
        alternatingPlain.Write(out);
        lifecycle.Write(out);
        manyPlain.Write(out);
        manySkipped.Write(out);

        out << L"reusedShapeIds=" << reusedIds << L";skipsAfterRecreate=" << wrongSkips << L';';
        out << L"manyPerShapePlainMs="
            << (manyPlain.MeanMs() / static_cast<double>(shapesInPass)) << L';'
            << L"manyPerShapeIfChangedMs="
            << (manySkipped.MeanMs() / static_cast<double>(shapesInPass)) << L';';
        if (hot.MeanMs() > 0.0) {
            out << L"skipSpeedup=" << (baseline.MeanMs() / hot.MeanMs()) << L';';
        }
        out << L"skipSavesMs=" << (baseline.MeanMs() - hot.MeanMs()) << L';'
            << L"ifChangedOverheadMs="
            << (alternatingIfChanged.MeanMs() - alternatingPlain.MeanMs()) << L';';
    } catch (...) {
        if (red) {
            BB_ReleaseTexture(red);
        }
        if (blue) {
            BB_ReleaseTexture(blue);
        }
        for (bb::Value& shape : created) {
            try {
                bb::call(shape.obj(), L"Delete");
            } catch (...) {
                // Already gone: the lifecycle leg deletes as it goes.
            }
        }
        throw;
    }

    BB_ReleaseTexture(red);
    BB_ReleaseTexture(blue);
    for (bb::Value& shape : created) {
        try {
            bb::call(shape.obj(), L"Delete");
        } catch (...) {
        }
    }
    return out.str();
}
