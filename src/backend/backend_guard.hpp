#pragma once
/**
 * @file backend_guard.hpp
 * The checks every Windows backend performs before it touches a caller's Shape,
 * and the exception-to-BackendResult boundary that sits around all of them.
 *
 * ## Why this is shared rather than written twice
 *
 * There are two Windows backends - the accelerated one that drives Office
 * internals on x64, and the portable one that drives documented Automation
 * everywhere else - and they must agree exactly about what a usable Shape is.
 * Two copies of "is this pointer a live IDispatch" would be two chances to
 * diverge, and the divergence would show up as one backend accepting a pointer
 * the other rejects: a difference a caller would experience as "BlipBridge works
 * on 64-bit and not on 32-bit", with no way to tell why.
 *
 * So the pointer validation, the class gate and the error translation live here
 * once. What the two backends do *after* the checks pass is where they differ,
 * and that is the only place they should.
 */

#include "backend.hpp"
#include "windows_office/shape_policy.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <new>
#include <stdexcept>

// MinGW requires Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

namespace bb {

/**
 * Maps an internal HRESULT onto the backend vocabulary.
 *
 * The distinction that matters to a caller is "your Office build is not
 * supported" versus "you passed something wrong", so E_NOTIMPL - which the
 * guards use for every unvalidated layout - becomes UnsupportedBuild.
 */
inline BackendStatus StatusFor(HRESULT hr) {
    switch (hr) {
    case E_INVALIDARG:
        return BackendStatus::InvalidArgument;
    case E_ACCESSDENIED:
        return BackendStatus::UnsupportedHost;
    case E_NOTIMPL:
        return BackendStatus::UnsupportedBuild;
    case E_OUTOFMEMORY:
        return BackendStatus::OutOfMemory;
    case RPC_E_WRONG_THREAD:
        return BackendStatus::Internal;
    default:
        break;
    }
    if (hr == BB_E_INVALID_IMAGE) {
        return BackendStatus::DecodeFailed;
    }
    if (hr == BB_E_TEXTURE_NOT_FOUND) {
        return BackendStatus::InvalidHandle;
    }
    if (hr == BB_E_SHAPE_CLASS_UNSUPPORTED) {
        return BackendStatus::UnsupportedShapeClass;
    }
    if (hr == BB_E_IMAGE_FILE_MISSING) {
        return BackendStatus::FileNotFound;
    }
    if (hr == BB_E_FALLBACK_REFUSED) {
        return BackendStatus::FallbackFailed;
    }
    return BackendStatus::Internal;
}

/// Runs @p body, converting any failure into a BackendResult.
template <typename Body>
BackendResult Guarded(Body&& body) noexcept {
    try {
        body();
        return BackendResult::Success();
    } catch (const Error& error) {
        return BackendResult::Failure(StatusFor(error.hr), error.what());
    } catch (const std::bad_alloc&) {
        return BackendResult::Failure(BackendStatus::OutOfMemory, "Out of memory");
    } catch (const std::exception& error) {
        return BackendResult::Failure(BackendStatus::Internal, error.what());
    } catch (...) {
        return BackendResult::Failure(BackendStatus::Internal, "Unknown internal failure");
    }
}

/// Releases an IDispatch on scope exit, so every early return stays correct.
struct ReleaseDispatch {
    IDispatch* value;

    ~ReleaseDispatch() {
        if (value) {
            value->Release();
        }
    }
};

/**
 * Validates that @p shape is a live COM object, and nothing more.
 *
 * VBA hands over `ObjPtr(shape)`, a raw IDispatch with no reference taken, so
 * this validates rather than trusts: the pointer must be readable and must
 * answer QueryInterface for IDispatch.
 *
 * Separate from RequireFillableShape because the picture path needs the pointer
 * checks without the class decision: a class the native path refuses may still
 * be fillable through Office's own API, and that choice is made further in.
 *
 * The returned pointer carries a reference the caller must release.
 */
inline IDispatch* RequireDispatchShape(void* shape) {
    if (!shape) {
        throw Error(E_INVALIDARG, "Shape pointer is null");
    }
    MEMORY_BASIC_INFORMATION information{};
    if (VirtualQuery(shape, &information, sizeof(information)) != sizeof(information) ||
        information.State != MEM_COMMIT) {
        throw Error(E_INVALIDARG, "Shape pointer does not address committed memory");
    }
    auto candidate = static_cast<IUnknown*>(shape);
    IDispatch* dispatch = nullptr;
    if (FAILED(candidate->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&dispatch))) ||
        !dispatch) {
        throw Error(E_INVALIDARG, "Shape pointer is not an IDispatch");
    }
    return dispatch;
}

/**
 * The same pointer checks, plus the semantic class gate.
 *
 * One authority decides which Shape classes are acceptable - see
 * shape_policy.hpp - so the C ABI, the COM surface and both backends cannot
 * disagree about what they accept. On the accelerated backend this is the gate
 * in front of the private OART apply, and nothing internal has been touched when
 * it refuses.
 *
 * @p shapeType receives `Shape.Type`, so a caller that needs it afterwards does
 * not pay for a second Automation fetch.
 */
inline IDispatch* RequireFillableShape(void* shape, long* shapeType = nullptr) {
    IDispatch* dispatch = RequireDispatchShape(shape);
    // The QueryInterface reference is only needed while the type is checked;
    // the caller's own reference keeps the Shape alive for the call.
    ReleaseDispatch release{dispatch};
    const office::ShapeClassification classification =
        office::RequireNativePictureFillTarget(dispatch);
    if (shapeType) {
        *shapeType = classification.shapeType;
    }
    return dispatch;
}

} // namespace bb
