/**
 * @file range_texture.cpp
 * The multi-Shape apply. See range_texture.hpp for why it takes a ShapeRange.
 */

#include "range_texture.hpp"

#include "apply_skip_cache.hpp"
#include "native_apply.hpp"
#include "native_texture.hpp"
#include "oart_layout.hpp"
#include "shape_identity.hpp"
#include "shape_policy.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <sstream>
#include <vector>

namespace bb::office {
namespace {

/// How many members @p range holds, with a legible refusal for a non-range.
long MemberCount(IDispatch* range) {
    try {
        return bb::get(range, L"Count").integer();
    } catch (const bb::Error&) {
        // A Shape has no Count. Saying so beats "member 1 is refused".
        throw bb::Error(E_INVALIDARG,
                        "This object is not a ShapeRange: it did not answer Count. Pass "
                        "Slide.Shapes.Range(...) or ShapeRange, not a single Shape - "
                        "BB_ApplyTexture takes a Shape.");
    }
}

/**
 * Classifies every member, refusing the whole range if any one is ineligible.
 *
 * Returns each member's `Shape.Type` in order, so the caller can key them
 * afterwards without asking Office for it a second time.
 *
 * This runs to completion before any internal object is touched. The failure it
 * prevents is not an error message: a native picture fill on a Connector
 * terminates PowerPoint, and the fill reaches every member of the range.
 */
std::vector<long> RequireEveryMemberEligible(IDispatch* range) {
    const long count = MemberCount(range);
    if (count <= 0) {
        throw bb::Error(E_INVALIDARG, "The ShapeRange holds no Shapes");
    }

    std::vector<long> types;
    types.reserve(static_cast<std::size_t>(count));
    for (long index = 1; index <= count; ++index) {
        bb::Value member = bb::call(range, L"Item", {bb::Value(index)});
        if (member.v.vt != VT_DISPATCH || !member.obj()) {
            std::ostringstream out;
            out << "Range member " << index << " of " << count << " is not a Shape";
            throw bb::Error(E_INVALIDARG, out.str());
        }
        const ShapeClassification verdict =
            ClassifyShapeForNativePictureFill(member.obj());
        if (!verdict.native()) {
            std::ostringstream out;
            out << "Range member " << index << " of " << count << " is refused: "
                << verdict.reason
                << ". The fill would reach every member, so the whole range is refused and "
                   "nothing has been applied.";
            throw bb::Error(verdict.eligibility == ShapeEligibility::Invalid
                                ? E_INVALIDARG
                                : bb::BB_E_SHAPE_CLASS_UNSUPPORTED,
                            out.str());
        }
        types.push_back(verdict.shapeType);
    }
    return types;
}

/**
 * Records what every member now carries, so a later skip cannot be wrong.
 *
 * A group and anything inside one is left unrecorded - `DescribeShape` refuses
 * to key them, because filling a group changes what its children render and
 * filling a child changes what the group shows. Neither side's remembered image
 * would stay true, so neither is remembered and both always do real work.
 */
void RememberEveryMember(IDispatch* range,
                         const std::vector<long>& types,
                         std::uint64_t textureId) noexcept {
    for (std::size_t index = 0; index < types.size(); ++index) {
        try {
            bb::Value member =
                bb::call(range, L"Item", {bb::Value(static_cast<long>(index) + 1)});
            if (member.v.vt != VT_DISPATCH || !member.obj()) {
                continue;
            }
            RememberApplied(DescribeShape(member.obj(), types[index]), textureId);
        } catch (...) {
            // Failing to record costs one apply next time. It must never fail
            // the apply that has already succeeded.
        }
    }
}

} // namespace

std::uint32_t ApplyTextureToRange(IDispatch* range, long handle) {
    if (!range) {
        throw bb::Error(E_POINTER, "Missing ShapeRange");
    }

    // Resolved first so an unknown handle fails exactly as BB_ApplyTexture's
    // would, before any Shape is questioned.
    const TextureRef texture = LookupTexture(handle);
    void* cached = CachedImageOf(texture);
    if (!cached) {
        throw bb::Error(bb::BB_E_TEXTURE_NOT_FOUND, "No image behind that texture handle");
    }

    const std::vector<long> types = RequireEveryMemberEligible(range);

    bb::Value fill = bb::get(range, L"Fill");
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        throw bb::Error(E_INVALIDARG, "The ShapeRange has no Fill object");
    }
    // Re-resolved for this apply and never cached: a range whose Shapes went
    // away still presents a structurally plausible pointer chain.
    const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill.obj());
    const bb::oart::ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);
    bb::oart::ApplyCachedImage(functions, target, cached);

    // One private call, but it filled every member, and the reuse diagnostic
    // counts fills rather than calls.
    CountTextureApplies(texture, static_cast<unsigned long>(types.size()));
    RememberEveryMember(range, types, TextureIdOf(texture));
    return static_cast<std::uint32_t>(types.size());
}

} // namespace bb::office
