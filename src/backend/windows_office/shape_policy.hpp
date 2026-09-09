#pragma once
/**
 * @file shape_policy.hpp
 * The one authority on whether a Shape may take the native picture-fill path.
 *
 * ## Structural validation is not semantic validation
 *
 * Structural validation - the walk in oart_layout.cpp - proves that the internal
 * objects behind a Shape are the Office objects we think they are: the PPCORE
 * delegating wrapper, the OART FillFormat, the control block, the receiver.
 *
 * Semantic validation asks a different question: is performing a picture fill on
 * *this class of Shape* a meaningful, safe operation at all?
 *
 * The two are independent, and a Connector proves it. A Connector presents the
 * identical wrapper, the identical FillFormat vtable, an identical control block
 * and an identical receiver as an ordinary rectangle - every structural check
 * passes - and it reports `Shape.Type = 1`, msoAutoShape. Office's own
 * `Fill.UserPicture` refuses it, from a check PPCORE performs *before* the fill
 * handler is reached. Our native apply reproduces the handler and not that
 * pre-check, and applying to one terminates PowerPoint.
 *
 * So `Shape.Type` is descriptive metadata, never proof. Connector and WordArt
 * both report msoAutoShape and behave completely differently. This header exists
 * so that judgement is made in exactly one place, from evidence, and every entry
 * point asks it rather than re-deriving it.
 *
 * See docs/shape_compatibility.md for the matrix and how a class earns a place
 * in it, and docs/safety_model.md for the two-layer model.
 */

#include <windows.h>
#include <oleauto.h>

#include <string>

namespace bb::office {

/// What may be done with a Shape, decided from its class rather than its layout.
enum class ShapeEligibility {
    /**
     * A validated native picture-fill path exists for this class: it was applied
     * to, verified, and survived the stress suite on the supported build.
     */
    NativeSupported,
    /**
     * No native path, but Office's own `Fill.UserPicture` accepts the class.
     * This is the only condition under which falling back is correct.
     */
    FallbackSupported,
    /**
     * A real Shape with no meaningful picture fill by either route - including
     * the classes that are *dangerous* natively, such as Connector and Line.
     */
    Unsupported,
    /**
     * Not a usable Shape: the object could not be questioned at all. This is a
     * failure to report, never a reason to try something slower.
     */
    Invalid,
};

/// The verdict plus the evidence it was reached from.
struct ShapeClassification {
    ShapeEligibility eligibility = ShapeEligibility::Invalid;
    /// `Shape.Type`, recorded for diagnostics. Never used alone as proof.
    long shapeType = 0;
    /// `Shape.Connector` as Office reports it; the discriminator no layout has.
    bool connector = false;
    /// A sentence explaining the verdict, suitable for the caller's error text.
    std::string reason;

    bool native() const { return eligibility == ShapeEligibility::NativeSupported; }
};

/**
 * Classifies @p shape for the native picture-fill path.
 *
 * Never throws: every outcome, including an unusable object, is a verdict the
 * caller has to act on rather than an exception to catch. Costs two Automation
 * property fetches, about two microseconds, against roughly 190 for an apply.
 *
 * @p shape is borrowed for the call and never retained.
 */
ShapeClassification ClassifyShapeForNativePictureFill(IDispatch* shape) noexcept;

/**
 * Throws bb::Error unless @p shape is `NativeSupported`.
 *
 * The gate in front of the private OART apply. `Unsupported` and `Invalid` carry
 * different HRESULTs so the C ABI can report "this Shape class has no native
 * path" separately from "this is not a usable Shape".
 */
void RequireNativePictureFillTarget(IDispatch* shape);

} // namespace bb::office
