/**
 * @file shape_compatibility.cpp
 * Classifies one Shape against the validated fill machinery, without calling
 * anything private and without touching the document.
 *
 * The backend accepts AutoShape and Freeform because those are the two classes
 * whose receiver chain was proved. That is a deliberately conservative list, not
 * a claim that nothing else works - and the only honest way to widen it is to
 * look at what each other class actually presents.
 *
 * So this walks the same chain `ResolveFillTarget` walks, but **reports** each
 * step instead of throwing at the first failure. A matrix needs to know *which*
 * step a class fails at: "has no Fill at all" and "has a Fill whose inner object
 * is some other OART type" are different findings with different consequences.
 *
 * ## Why it double-checks itself
 *
 * The step-by-step walk here could drift away from the real guard in
 * oart_layout.cpp, and a compatibility matrix built on a drifted copy would be
 * worse than none. So after classifying, it calls the real `ResolveFillTarget`
 * and reports whether the two agree. A disagreement is reported as
 * `agrees=0`, which is a defect in this file, and the harness treats it as a
 * failed probe rather than a result.
 *
 * Nothing here is a decision about support. It produces evidence; the harness
 * `tools/test_shape_compatibility.ps1` decides, and only after an apply has been
 * attempted and verified.
 *
 * Read-only, borrows every pointer, retains nothing. STA, PowerPoint host.
 */

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../../src/backend/windows_office/shape_policy.hpp"
#include "../experiment_api.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <sstream>
#include <string>

namespace {

/// How much of a candidate FillFormat must be readable before it is dereferenced.
constexpr std::size_t kPublicFillProbeSize = 0x10;
/// How much of the inner OART object must be readable before its vtable is read.
constexpr std::size_t kInnerProbeSize = 0x68;

/**
 * The step at which a Shape stops matching the validated chain.
 *
 * Ordered: each value means every earlier step succeeded. `Complete` is the only
 * value that says the chain is intact, and even that is not yet a decision to
 * support the class - an apply still has to be attempted and verified.
 */
enum class Step {
    NoShapeType,        ///< Shape.Type could not be read; not a usable Shape.
    NoFillProperty,     ///< The object has no Fill at all.
    FillNotReadable,    ///< Fill returned something that is not committed memory.
    NotWrapper,         ///< Fill is not a PPCORE delegating wrapper.
    InnerNotReadable,   ///< The wrapper's inner pointer is not committed memory.
    InnerNotFillFormat, ///< The inner object is some other OART type.
    NoReceiver,         ///< The control block or receiver did not check out.
    Complete            ///< The full validated chain is present.
};

const wchar_t* StepName(Step step) {
    switch (step) {
    case Step::NoShapeType:
        return L"NoShapeType";
    case Step::NoFillProperty:
        return L"NoFillProperty";
    case Step::FillNotReadable:
        return L"FillNotReadable";
    case Step::NotWrapper:
        return L"NotWrapper";
    case Step::InnerNotReadable:
        return L"InnerNotReadable";
    case Step::InnerNotFillFormat:
        return L"InnerNotFillFormat";
    case Step::NoReceiver:
        return L"NoReceiver";
    case Step::Complete:
        return L"Complete";
    }
    return L"Unknown";
}

/// Escapes the separators the report format reserves, so a COM error message
/// containing one cannot forge a field.
std::wstring Sanitise(const std::string& text) {
    std::wstring out;
    out.reserve(text.size());
    for (unsigned char character : text) {
        if (character == ';' || character == '=' || character < 0x20) {
            out.push_back(L' ');
        } else if (character < 0x80) {
            out.push_back(static_cast<wchar_t>(character));
        } else {
            out.push_back(L'?');
        }
    }
    return out;
}

} // namespace

/**
 * Reports how far @p shape gets along the validated fill chain.
 *
 * @p shape is a PowerPoint Shape, not a FillFormat: the point is to classify a
 * Shape category, and getting `Fill` is itself one of the steps that can fail.
 */
std::wstring probeShapeCompatibility(IDispatch* shape) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Shape probing requires the PowerPoint host");
    }

    std::wostringstream out;
    out << L"build=" << bb::oart::kSupportedVersionText << L';';

    Step reached = Step::NoShapeType;
    long shapeType = 0;
    try {
        shapeType = bb::get(shape, L"Type").integer();
        out << L"shapeType=" << shapeType << L';';
        reached = Step::NoFillProperty;
    } catch (const bb::Error& error) {
        out << L"shapeType=?;error=" << Sanitise(error.what()) << L';' << L"step="
            << StepName(reached) << L";agrees=1;";
        return out.str();
    }

    // Some classes raise rather than return when asked for a fill they do not
    // have. That is a result, not an error, so it is caught and named.
    bb::Value fill;
    try {
        fill = bb::get(shape, L"Fill");
    } catch (const bb::Error& error) {
        out << L"fillError=" << Sanitise(error.what()) << L";step=" << StepName(reached)
            << L";agrees=1;";
        return out.str();
    }
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        out << L"fillError=Fill is not an object;step=" << StepName(reached) << L";agrees=1;";
        return out.str();
    }
    IDispatch* const fillObject = fill.obj();
    out << L"hasFill=1;";

    // From here the walk mirrors ResolveFillTarget step for step, reporting
    // rather than throwing. The agreement check at the end is what keeps the
    // two honest about each other.
    const bb::oart::OfficeModules modules = bb::oart::SynchroniseOfficeModules();
    if (!modules.oart || !modules.ppcore) {
        throw bb::Error(E_NOTIMPL, "OART and PPCORE must both be loaded to classify a Shape");
    }
    const auto oartBase = reinterpret_cast<std::uintptr_t>(modules.oart);
    const auto ppcoreBase = reinterpret_cast<std::uintptr_t>(modules.ppcore);

    reached = Step::FillNotReadable;
    if (bb::oart::IsReadable(fillObject, kPublicFillProbeSize)) {
        reached = Step::NotWrapper;
        bb::oart::DelegatingWrapper wrapper;
        const std::size_t ppcoreSize = bb::oart::OfficeModuleImageSize(modules.ppcore);
        if (ppcoreSize != 0 &&
            bb::oart::DescribeDelegatingWrapper(fillObject, ppcoreBase, ppcoreSize, wrapper)) {
            out << L"wrapperVtable=ppcore.dll+0x" << std::hex << wrapper.vtableRva << std::dec
                << L";wrapperInnerOffset=0x" << std::hex << wrapper.innerOffset << std::dec
                << L";wrapperThunks=" << wrapper.identityThunks << L';';

            reached = Step::InnerNotReadable;
            auto* inner =
                reinterpret_cast<void*>(bb::oart::LoadPointer(fillObject, wrapper.innerOffset));
            if (bb::oart::IsReadable(inner, kInnerProbeSize)) {
                const std::uintptr_t innerVtable = bb::oart::LoadPointer(inner, 0);
                out << L"innerVtable=0x" << std::hex << innerVtable << std::dec << L';';
                if (innerVtable >= oartBase) {
                    out << L"innerVtableRva=oart.dll+0x" << std::hex << (innerVtable - oartBase)
                        << std::dec << L';';
                }
                reached = innerVtable == oartBase + bb::oart::kFillFormatVtableRva
                              ? Step::NoReceiver
                              : Step::InnerNotFillFormat;
            }
        } else {
            // Still worth recording what it *is*, so a class that turns out to
            // wrap something else is identifiable rather than just "not a
            // wrapper".
            out << L"fillVtable=0x" << std::hex << bb::oart::LoadPointer(fillObject, 0) << std::dec
                << L';';
        }
    }

    // This file's own conclusion, recorded *before* the real guard runs, or the
    // comparison below would be comparing the guard against itself.
    //
    // The walk above stops once the inner object presents the FillFormat vtable;
    // it does not check the control block or the receiver. So it cannot predict
    // success - only failure. The invariant it can assert is one-directional:
    //
    //     this walk failed early  =>  the guard must also refuse
    //
    // The converse is allowed: reaching NoReceiver here means "nothing has ruled
    // it out yet", and the guard may still refuse at the control block.
    const bool couldStillResolve = reached == Step::NoReceiver;

    bool resolved = false;
    std::wstring guardMessage;
    try {
        const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
        resolved = true;
        out << L"receiver=0x" << std::hex << reinterpret_cast<std::uintptr_t>(target.receiver)
            << L";handler=0x" << reinterpret_cast<std::uintptr_t>(target.handler) << std::dec
            << L";tokenStrong=" << target.tokenStrong << L';';
        reached = Step::Complete;
    } catch (const bb::Error& error) {
        guardMessage = Sanitise(error.what());
    }

    // A violation means this classifier is wrong about the layout, so the
    // harness must discard the row rather than publish it as a matrix result.
    const bool agrees = couldStillResolve || !resolved;
    out << L"step=" << StepName(reached) << L";resolved=" << (resolved ? 1 : 0) << L";agrees="
        << (agrees ? 1 : 0) << L';';
    if (!guardMessage.empty()) {
        out << L"guard=" << guardMessage << L';';
    }
    return out.str();
}

/**
 * Applies a texture to @p shape with the **Shape-type allowlist bypassed**.
 *
 * The matrix showed that TextBox, Placeholder, Callout, Group, Picture and Media
 * all present the identical validated wrapper, FillFormat and receiver as an
 * AutoShape, and are refused only because they are not on the allowlist. Whether
 * that means an apply is *safe* is a separate question, and the Connector
 * answered it emphatically: identical structure, and applying to one terminates
 * PowerPoint.
 *
 * So the allowlist cannot be widened by reasoning; it can only be widened by
 * trying, in a process whose death is an acceptable outcome. That is what this
 * is for, and it is why it is research-only and reachable solely through the
 * COM surface - never through the C ABI.
 *
 * Everything except the type allowlist still applies: the connector refusal
 * (which is not an allowlist entry but a known-fatal case), every module and
 * version guard, every vtable check, every signature check, and the receiver
 * re-resolution inside nativeTextureApply.
 */
std::wstring applyTextureUnrestricted(IDispatch* shape, long handle) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Applying requires the PowerPoint host");
    }

    // Connectors are excluded even here. Their fatality is established; running
    // that experiment again would only cost another PowerPoint.
    constexpr long kMsoTrue = -1;
    bool isConnector = true;
    try {
        isConnector = bb::get(shape, L"Connector").integer() == kMsoTrue;
    } catch (const bb::Error&) {
        throw bb::Error(E_INVALIDARG, "Shape would not say whether it is a connector");
    }
    if (isConnector) {
        throw bb::Error(E_INVALIDARG,
                        "Connectors and lines are known to terminate the host; refused "
                        "even in the unrestricted path");
    }

    bb::Value fill = bb::get(shape, L"Fill");
    if (fill.v.vt != VT_DISPATCH || !fill.obj()) {
        throw bb::Error(E_INVALIDARG, "Shape has no Fill object");
    }
    nativeTextureApply(fill.obj(), handle);

    std::wostringstream out;
    out << L"applied=1;allowlistBypassed=1;";
    return out.str();
}

/**
 * Reports the semantic verdict for @p shape and the private-apply entry count.
 *
 * Both in one call so a harness can read the count, attempt an apply, and read
 * it again without a third round trip changing anything in between.
 */
std::wstring probeShapePolicy(IDispatch* shape) {
    const bb::office::ShapeClassification classification =
        bb::office::ClassifyShapeForNativePictureFill(shape);
    const wchar_t* name = L"Invalid";
    switch (classification.eligibility) {
    case bb::office::ShapeEligibility::NativeSupported:
        name = L"NativeSupported";
        break;
    case bb::office::ShapeEligibility::FallbackSupported:
        name = L"FallbackSupported";
        break;
    case bb::office::ShapeEligibility::Unsupported:
        name = L"Unsupported";
        break;
    case bb::office::ShapeEligibility::Invalid:
        name = L"Invalid";
        break;
    }
    std::wostringstream out;
    out << L"eligibility=" << name << L";shapeType=" << classification.shapeType << L";connector="
        << (classification.connector ? 1 : 0) << L";applyEntries="
        << bb::oart::NativeApplyEntryCount() << L";reason=" << Sanitise(classification.reason)
        << L';';
    return out.str();
}
