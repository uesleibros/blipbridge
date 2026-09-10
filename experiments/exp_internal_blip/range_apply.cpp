/**
 * @file range_apply.cpp
 * One private apply that fills every Shape in a ShapeRange.
 *
 * ## Why this is the batch, and the earlier idea was not
 *
 * The transaction handed to the receiver carries no target: the receiver *is*
 * the target. So a transaction holding several picture fills for several Shapes
 * cannot exist, and looking for one was looking in the wrong place. What can
 * exist is a receiver that stands for several Shapes - and that is exactly what
 * `ShapeRange.Fill` resolves to. It is the same PPCORE delegating wrapper as
 * `Shape.Fill`, over the same OART FillFormat, over a control block whose
 * receiver carries the whole range, so the existing structural walk accepts it
 * unchanged and one apply fills all of them.
 *
 * Measured on 16.0.14334.20848, 200 rounds x 3 runs, gate and per-Shape record
 * included on both sides, both legs in this process:
 *
 *                 range      one at a time
 *      1 Shape    0.232 ms   0.229 ms
 *      2 Shapes   0.299 ms   0.449 ms
 *      8 Shapes   0.632 ms   1.771 ms
 *     32 Shapes   1.876 ms   6.650 ms      3.5x
 *    100 Shapes   5.549 ms  20.935 ms      3.8x
 *
 * so 0.059 ms per Shape at 32 against 0.208, and no gain at all at one - the
 * range machinery has its own fixed cost, and only sharing it pays.
 *
 * ## Why this is not simply better
 *
 * **It makes no undo entry.** Measured, not assumed: five Undos leave a
 * range-filled Shape filled, while one Undo reverts a single-Shape native apply
 * in the same document, and one Undo reverts Office's own
 * ShapeRange.Fill.UserPicture across every member. So Office does record undo
 * for range fills, above the receiver this reaches, and going straight to the
 * receiver skips it.
 *
 * That is a semantic difference, not a tuning detail, and it is why this cannot
 * quietly become what ApplyTexture does. The fills persist through save and
 * reopen and every member is filled correctly - see tools/test_range_apply.ps1 -
 * but a caller who fills 32 Shapes and presses Ctrl+Z gets nothing back.
 *
 * ## The part that is not free
 *
 * A ShapeRange is a bag of whatever the caller put in it, and the fill goes to
 * every member. A Connector in that bag would reach the private backend, and a
 * native picture fill on a Connector terminates PowerPoint - so **every member
 * is classified before anything internal is touched**, and one ineligible member
 * refuses the whole range by name. That check costs two Automation reads per
 * member, which is why the timing above is the apply alone and the numbers this
 * file reports include the gate.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/apply_skip_cache.hpp"
#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../../src/backend/windows_office/shape_identity.hpp"
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

/**
 * Classifies every member of @p range, refusing the whole range if any member
 * has no validated native path.
 *
 * Returns each member's `Shape.Type`, in order, so the caller can key them
 * afterwards without asking Office a second time.
 *
 * This is the gate. It runs to completion before any internal object is touched,
 * because the failure it prevents is not an error message - it is PowerPoint
 * closing with the user's document in it.
 */
std::vector<long> RequireEveryMemberEligible(IDispatch* range) {
    const long count = bb::get(range, L"Count").integer();
    if (count <= 0) {
        throw bb::Error(E_INVALIDARG, "The range holds no Shapes");
    }

    std::vector<long> types;
    types.reserve(static_cast<std::size_t>(count));
    for (long index = 1; index <= count; ++index) {
        bb::Value member = bb::call(range, L"Item", {bb::Value(index)});
        if (member.v.vt != VT_DISPATCH || !member.obj()) {
            throw bb::Error(E_INVALIDARG, "A range member is not a Shape");
        }
        const bb::office::ShapeClassification verdict =
            bb::office::ClassifyShapeForNativePictureFill(member.obj());
        if (!verdict.native()) {
            std::ostringstream out;
            out << "Range member " << index << " of " << count << " is refused: " << verdict.reason
                << ". The fill would reach every member, so the range is refused whole.";
            throw bb::Error(verdict.eligibility == bb::office::ShapeEligibility::Invalid
                                ? E_INVALIDARG
                                : bb::BB_E_SHAPE_CLASS_UNSUPPORTED,
                            out.str());
        }
        types.push_back(verdict.shapeType);
    }
    return types;
}

/// Records what every member now carries, so a later skip cannot be wrong.
void RememberEveryMember(IDispatch* range,
                         const std::vector<long>& types,
                         std::uint64_t textureId) {
    for (std::size_t index = 0; index < types.size(); ++index) {
        bb::Value member = bb::call(range, L"Item", {bb::Value(static_cast<long>(index) + 1)});
        if (member.v.vt != VT_DISPATCH || !member.obj()) {
            continue;
        }
        bb::office::RememberApplied(bb::office::DescribeShape(member.obj(), types[index]),
                                    textureId);
    }
}

} // namespace

std::wstring applyTextureToRange(IDispatch* range, long handle, long iterations) {
    if (!range) {
        throw bb::Error(E_POINTER, "Missing ShapeRange");
    }
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }

    const bb::office::TextureRef texture = bb::office::LookupTexture(handle);
    void* cached = bb::office::CachedImageOf(texture);
    if (!cached) {
        throw bb::Error(E_FAIL, "No image behind that handle");
    }
    const std::uint64_t textureId = bb::office::TextureIdOf(texture);

    const double tick = SecondsPerTick();
    std::vector<double> gateSamples;
    std::vector<double> applySamples;
    std::vector<double> totalSamples;
    std::vector<double> loopSamples;
    gateSamples.reserve(static_cast<std::size_t>(iterations));
    applySamples.reserve(static_cast<std::size_t>(iterations));
    totalSamples.reserve(static_cast<std::size_t>(iterations));
    loopSamples.reserve(static_cast<std::size_t>(iterations));

    // Warm every cache the path uses, so the samples are the steady state.
    std::vector<long> warmTypes = RequireEveryMemberEligible(range);
    {
        bb::Value fill = bb::get(range, L"Fill");
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
        bb::oart::ApplyCachedImage(
            bb::oart::ResolveApplyFunctions(target.oartBase), target, cached);
    }

    for (long round = 0; round < iterations; ++round) {
        const long long start = Now();
        const std::vector<long> types = RequireEveryMemberEligible(range);
        const long long gated = Now();

        bb::Value fill = bb::get(range, L"Fill");
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
        const bb::oart::ApplyFunctions functions =
            bb::oart::ResolveApplyFunctions(target.oartBase);
        bb::oart::ApplyCachedImage(functions, target, cached);
        const long long applied = Now();

        RememberEveryMember(range, types, textureId);
        const long long done = Now();

        gateSamples.push_back(static_cast<double>(gated - start) * tick * 1000.0);
        applySamples.push_back(static_cast<double>(applied - gated) * tick * 1000.0);
        totalSamples.push_back(static_cast<double>(done - start) * tick * 1000.0);

        /*
         * The same fills, one Shape at a time, in this process.
         *
         * It has to be measured here rather than from the harness: driving it
         * from PowerShell adds a cross-process Automation round trip per Shape,
         * which is about 5.9 ms and would swamp the 0.19 ms being compared. A
         * comparison like that would flatter the range by twenty-five times for
         * reasons that have nothing to do with Office.
         */
        const long long loopStart = Now();
        for (long index = 1; index <= static_cast<long>(types.size()); ++index) {
            bb::Value member = bb::call(range, L"Item", {bb::Value(index)});
            if (member.v.vt != VT_DISPATCH || !member.obj()) {
                continue;
            }
            bb::office::RequireNativePictureFillTarget(member.obj());
            bb::Value memberFill = bb::get(member.obj(), L"Fill");
            const bb::oart::FillTarget memberTarget =
                bb::oart::ResolveFillTarget(memberFill.obj());
            bb::oart::ApplyCachedImage(
                bb::oart::ResolveApplyFunctions(memberTarget.oartBase), memberTarget, cached);
            bb::office::RememberApplied(
                bb::office::DescribeShape(member.obj(), types[static_cast<std::size_t>(index) - 1]),
                textureId);
        }
        loopSamples.push_back(static_cast<double>(Now() - loopStart) * tick * 1000.0);
    }

    const auto summarise = [&](std::vector<double>& samples, const wchar_t* name,
                               std::wostringstream& out) {
        std::sort(samples.begin(), samples.end());
        double total = 0.0;
        for (double sample : samples) {
            total += sample;
        }
        const double mean = total / static_cast<double>(samples.size());
        out << name << L"MeanMs=" << mean << L';' << name << L"MedianMs="
            << samples[samples.size() / 2] << L';';
        return mean;
    };

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    const auto shapes = static_cast<double>(warmTypes.size());
    out << L"iterations=" << iterations << L";shapes=" << warmTypes.size() << L';';
    const double gateMean = summarise(gateSamples, L"gate", out);
    const double applyMean = summarise(applySamples, L"apply", out);
    const double totalMean = summarise(totalSamples, L"total", out);
    const double loopMean = summarise(loopSamples, L"loop", out);
    out << L"perShapeApplyMs=" << (applyMean / shapes) << L';' << L"perShapeTotalMs="
        << (totalMean / shapes) << L';' << L"perShapeLoopMs=" << (loopMean / shapes) << L';'
        << L"gateShare=" << (totalMean > 0.0 ? 100.0 * gateMean / totalMean : 0.0) << L';'
        << L"speedup=" << (totalMean > 0.0 ? loopMean / totalMean : 0.0) << L';';
    return out.str();
}
