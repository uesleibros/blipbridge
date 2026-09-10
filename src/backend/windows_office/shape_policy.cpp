/**
 * @file shape_policy.cpp
 * The semantic eligibility table, and the gate that reads it.
 *
 * Every entry in the tables below was earned by running a real instance of the
 * class through `tools/test_shape_compatibility.ps1` and, for the native ones,
 * `tools/test_shape_class_stress.ps1`. None of it was inferred from the fact
 * that a class exposes a `Fill` property, and none of it may be.
 *
 * The tables are a property of **one Office build**, 16.0.14334.20848. On any
 * other build the structural guards refuse first, so an unvalidated build can
 * never reach these decisions; re-running the harnesses is what would earn the
 * right to widen them.
 */

#include "shape_policy.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstddef>
#include <sstream>

namespace bb::office {
namespace {

/// `Shape.Connector`, like every Office boolean, is msoTrue = -1.
constexpr long kMsoTrue = -1;

/**
 * Classes with a validated native picture-fill path, by msoShapeType.
 *
 * Each one presents the validated wrapper, FillFormat and receiver; a native
 * apply leaves identity, type, geometry and rotation untouched and produces a
 * picture fill; and it survived 1000 repeated applies, 500 alternating applies,
 * undo, redo, save, reopen, deletion and presentation close with the cached
 * image's reference count returning to the handle's own.
 *
 * WordArt and group children are not listed separately because they report
 * msoAutoShape - which is exactly why the connector check below exists.
 */
constexpr long kNativeShapeTypes[] = {
    1,  // msoAutoShape - also WordArt and group children
    2,  // msoCallout
    5,  // msoFreeform
    6,  // msoGroup
    13, // msoPicture
    14, // msoPlaceholder
    16, // msoMedia
    17, // msoTextBox
};

/**
 * Classes with no native path, where Office's own `Fill.UserPicture` works.
 *
 * These are the *only* classes for which falling back is correct. Falling back
 * is a statement about the class, never a response to a failure.
 */
constexpr long kFallbackShapeTypes[] = {
    7,  // msoEmbeddedOLEObject
    19, // msoTable
    24, // msoIgraphic - what SmartArt reports on this build
    21, // msoDiagram - the older SmartArt reporting, refused the same way
};

/// Deduces the table's size, so adding a class cannot silently break the lookup.
template <std::size_t Count>
bool Contains(const long (&table)[Count], long value) {
    for (const long candidate : table) {
        if (candidate == value) {
            return true;
        }
    }
    return false;
}

} // namespace

ShapeClassification ClassifyShapeForNativePictureFill(IDispatch* shape) noexcept {
    ShapeClassification result;
    if (!shape) {
        result.eligibility = ShapeEligibility::Invalid;
        result.reason = "No Shape was given";
        return result;
    }

    try {
        result.shapeType = bb::get(shape, L"Type").integer();
    } catch (...) {
        // A Shape that will not report its own type is not a Shape we can reason
        // about. That is a failure to surface, not a reason to try the fallback.
        result.eligibility = ShapeEligibility::Invalid;
        result.reason = "The object did not answer Shape.Type, so it is not a usable Shape";
        return result;
    }

    /*
     * The connector question comes before the type table, because a Connector
     * reports msoAutoShape and would otherwise pass it.
     *
     * A connector is a line: it has no interior to fill. Office's own
     * Fill.UserPicture refuses one with "value out of range", from a check
     * PPCORE performs before the fill handler runs. The native apply reproduces
     * the handler and not that pre-check, and applying to one **terminates
     * PowerPoint** - reproduced in tools/test_connector_isolation.ps1.
     *
     * `Shape.Connector` is msoTrue for exactly Connector and Line and msoFalse
     * for every other class tested. It is Office's own answer, and it is the
     * only thing that separates a connector from a rectangle: every structural
     * check sees identical objects.
     *
     * A Shape that will not answer is refused rather than assumed safe. The cost
     * of being wrong here is the user's document.
     */
    try {
        result.connector = bb::get(shape, L"Connector").integer() == kMsoTrue;
    } catch (...) {
        result.eligibility = ShapeEligibility::Invalid;
        result.reason = "The Shape did not answer whether it is a connector, so it is refused";
        return result;
    }
    if (result.connector) {
        result.eligibility = ShapeEligibility::Unsupported;
        result.reason = "Connectors and lines have no fillable interior. Office's own "
                        "Fill.UserPicture refuses them, and a native apply terminates PowerPoint, "
                        "so this Shape is refused before any internal call is made";
        return result;
    }

    if (Contains(kNativeShapeTypes, result.shapeType)) {
        result.eligibility = ShapeEligibility::NativeSupported;
        result.reason = "This Shape class has a validated native picture-fill path";
        return result;
    }
    if (Contains(kFallbackShapeTypes, result.shapeType)) {
        result.eligibility = ShapeEligibility::FallbackSupported;
        result.reason = "This Shape class has no validated native picture-fill path, but Office's "
                        "own Fill.UserPicture accepts it";
        return result;
    }

    std::ostringstream out;
    out << "Shape type " << result.shapeType
        << " has no picture-fill path, native or otherwise, on this Office build; "
           "see docs/shape_compatibility.md";
    result.eligibility = ShapeEligibility::Unsupported;
    result.reason = out.str();
    return result;
}

ShapeClassification RequireNativePictureFillTarget(IDispatch* shape) {
    const ShapeClassification classification = ClassifyShapeForNativePictureFill(shape);
    switch (classification.eligibility) {
    case ShapeEligibility::NativeSupported:
        return classification;
    case ShapeEligibility::Invalid:
        throw bb::Error(E_INVALIDARG, classification.reason);
    case ShapeEligibility::FallbackSupported:
    case ShapeEligibility::Unsupported:
        // A real Shape, refused because of what it is rather than because
        // something went wrong with it. The C ABI reports this separately.
        throw bb::Error(bb::BB_E_SHAPE_CLASS_UNSUPPORTED, classification.reason);
    }
    throw bb::Error(E_INVALIDARG, "Unclassified Shape");
}

} // namespace bb::office
