/**
 * @file native_apply_probe.cpp
 * The two research entry points that drive the native apply directly.
 *
 * These predate the texture handle: they create a cached image, apply it, and
 * release it within the call, reporting the reference count at every stage so a
 * delta can be explained rather than assumed. `test_native_apply.ps1` and
 * `test_native_apply_reuse.ps1` are their harnesses.
 *
 * They live here rather than beside ApplyCachedImage because they are research:
 * nothing in the shipping path calls them, and the production backend should not
 * have to compile against the research header to define them.
 *
 * Research only. STA, PowerPoint host. Nothing is retained after the call.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <sstream>
#include <string>
#include <utility>
#include <vector>

using bb::oart::ApplyFunctions;
using bb::oart::FillTarget;

namespace {

/**
 * Owns one intrusive reference for the duration of a probe.
 *
 * The production apply has its own equivalent; this is deliberately a separate
 * three lines rather than exported internals, so that a research probe can never
 * be the reason something in the shipping path has to become public.
 */
class OwnedReference {
public:
    OwnedReference() = default;
    OwnedReference(const OwnedReference&) = delete;
    OwnedReference& operator=(const OwnedReference&) = delete;
    ~OwnedReference() { bb::oart::ReleaseIntrusive(value); }

    void* value = nullptr;
};


/// Records the cached image's count at each stage so deltas can be explained.
class CountTimeline {
public:
    explicit CountTimeline(const void* object) : object_(object) {}

    /// Distinguishes one Shape's samples from the next when reusing an image.
    void SetGroup(long group) { group_ = group; }

    void Sample(const wchar_t* label) {
        std::wstring name;
        if (group_ >= 0) {
            name = L"s" + std::to_wstring(group_) + L'.';
        }
        name += label;
        samples_.emplace_back(std::move(name), bb::oart::IntrusiveCount(object_));
    }

    void Append(std::wostringstream& out) const {
        for (const auto& [label, count] : samples_) {
            out << label << L'=' << count << L';';
        }
    }

private:
    const void* object_ = nullptr;
    long group_ = -1;
    std::vector<std::pair<std::wstring, std::uint32_t>> samples_;
};

} // namespace

/**
 * Applies @p bytes as a picture fill on the Shape behind @p fill, natively.
 *
 * Ownership: the cached image created here is released before returning. What
 * the document keeps is Office's own reference, taken when the record was
 * copied into the transaction and again during the apply. Nothing is retained
 * by BlipBridge, so this is deliberately not yet a texture handle.
 */
std::wstring nativeApplyExperiment(IDispatch* fill, SAFEARRAY* bytes) {
    // Resolved immediately before use and never stored: a deleted Shape still
    // passes every check in this chain.
    const FillTarget target = bb::oart::ResolveFillTarget(fill);
    const ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);

    const bb::oart::CreatedImage created = bb::oart::CreateCachedImageFromBytes(bytes);
    OwnedReference cached;
    OwnedReference image;
    cached.value = created.cached;
    image.value = created.image;

    CountTimeline timeline(cached.value);
    timeline.Sample(L"create");
    bb::oart::ApplyCachedImage(functions, target, cached.value,
                               [&](const wchar_t* label) { timeline.Sample(label); });

    std::wostringstream out;
    out << L"build=" << bb::oart::kSupportedVersionText << L';';
    out << L"receiver=0x" << std::hex << reinterpret_cast<std::uintptr_t>(target.receiver)
        << L";cached=0x" << reinterpret_cast<std::uintptr_t>(cached.value) << std::dec
        << L';';
    out << L"usedUserPicture=0;usedDonorShape=0;usedSourceFile=0;fromMemoryStream=1;";
    timeline.Append(out);
    return out.str();
}

/**
 * Applies **one** cached image to every FillFormat in @p fills.
 *
 * This is the reuse question the whole project turns on: the image is decoded
 * once, and each Shape then goes through record construction and the receiver
 * call only. The report names the single cached-image address and repeats it per
 * Shape, so reuse is visible rather than asserted.
 *
 * Each Shape's receiver is resolved immediately before its own apply and is
 * never cached between them.
 */
std::wstring nativeApplyReuseExperiment(SAFEARRAY* fills, SAFEARRAY* bytes) {
    if (!fills || SafeArrayGetDim(fills) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional array of FillFormats");
    }
    LONG lower = 0;
    LONG upper = 0;
    bb::check(SafeArrayGetLBound(fills, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(fills, 1, &upper), "SafeArrayGetUBound");
    if (upper < lower) {
        throw bb::Error(E_INVALIDARG, "No FillFormats supplied");
    }

    const bb::oart::CreatedImage created = bb::oart::CreateCachedImageFromBytes(bytes);
    OwnedReference cached;
    OwnedReference image;
    cached.value = created.cached;
    image.value = created.image;

    CountTimeline timeline(cached.value);
    timeline.Sample(L"create");

    std::wostringstream out;
    out << L"build=" << bb::oart::kSupportedVersionText << L";cached=0x" << std::hex
        << reinterpret_cast<std::uintptr_t>(cached.value) << std::dec
        << L";creations=1;shapes=" << (upper - lower + 1) << L';';

    VARTYPE elementType = VT_EMPTY;
    bb::check(SafeArrayGetVartype(fills, &elementType), "SafeArrayGetVartype");
    if (elementType != VT_VARIANT && elementType != VT_DISPATCH) {
        throw bb::Error(E_INVALIDARG, "FillFormat array must hold objects");
    }

    for (LONG index = lower; index <= upper; ++index) {
        timeline.SetGroup(index - lower);
        bb::Value element;
        if (elementType == VT_DISPATCH) {
            IDispatch* raw = nullptr;
            bb::check(SafeArrayGetElement(fills, &index, &raw), "SafeArrayGetElement");
            element = bb::Value(raw);
            if (raw) {
                raw->Release();   // SafeArrayGetElement already AddRef'd for us
            }
        } else {
            bb::check(SafeArrayGetElement(fills, &index, &element.v), "SafeArrayGetElement");
        }
        // Re-resolved per Shape; never cached across applies.
        const FillTarget target = bb::oart::ResolveFillTarget(element.obj());
        const ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);
        bb::oart::ApplyCachedImage(functions, target, cached.value,
                                   [&](const wchar_t* label) { timeline.Sample(label); });
        out << L"shape" << (index - lower) << L"Receiver=0x" << std::hex
            << reinterpret_cast<std::uintptr_t>(target.receiver) << std::dec << L';';
    }

    timeline.Append(out);
    return out.str();
}
