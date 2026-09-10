#pragma once
/**
 * @file shape_identity.hpp
 * Identifying a Shape well enough to remember something about it.
 *
 * Two caches need this and they must agree: `UserPicture2`'s per-Shape skip
 * state, and the explicit texture API's `ApplyTextureIfChanged`. Duplicating the
 * rules would let them drift, and the failure mode of a drifted Shape key is a
 * silently wrong picture rather than an error - so there is one implementation.
 *
 * ## What a key is, and what it deliberately is not
 *
 * It is a composite of *values read out of the object*: the presentation's
 * identity, the slide's `SlideID`, and the Shape's `Id`. It is **not** a pointer.
 * A pointer can be freed and its address reused, so a pointer key would
 * eventually name a different Shape without anything looking wrong. Values
 * cannot do that, and nothing here has to be told when a Shape dies.
 *
 * `Shape.Id` alone would not do either: it is unique within a slide, not across
 * a presentation, and certainly not across open documents.
 *
 * ## Groups are excluded, on evidence
 *
 * Filling a group changes what its children render, and filling a child changes
 * what the group shows - measured in `tools/probe_group_fill_propagation.ps1`,
 * where a child's rendered PNG went from 33,515 to 35,725 bytes with `Fill.Type`
 * reading 6 throughout. Either side's remembered texture therefore goes stale
 * when the other is filled, and no cheap property reveals it. A group and
 * anything inside one is left unkeyed, which costs one apply and never risks a
 * wrong one.
 *
 * Reading a key costs about four Automation property fetches - a few
 * microseconds against roughly 170 for the Office apply it may avoid.
 */

// MinGW requires the Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

namespace bb::office {

/**
 * A Shape's identity as values.
 *
 * `valid` false means this Shape must not be cached at all. Callers apply
 * normally in that case; they just pay for it.
 */
struct ShapeKey {
    long presentation = 0;
    long slide = 0;
    long shape = 0;
    bool valid = false;

    bool operator<(const ShapeKey& other) const {
        if (presentation != other.presentation) {
            return presentation < other.presentation;
        }
        if (slide != other.slide) {
            return slide < other.slide;
        }
        return shape < other.shape;
    }

    bool operator==(const ShapeKey& other) const {
        return valid && other.valid && presentation == other.presentation &&
               slide == other.slide && shape == other.shape;
    }
};

/**
 * Reads @p shape's composite identity.
 *
 * @p shapeType is `Shape.Type`, passed in because every caller has already read
 * it; fetching it again would be a second Automation call for nothing.
 *
 * Never throws. An unkeyable Shape is a reason to skip a cache, not to fail an
 * apply that would otherwise have worked.
 */
ShapeKey DescribeShape(IDispatch* shape, long shapeType) noexcept;

} // namespace bb::office
