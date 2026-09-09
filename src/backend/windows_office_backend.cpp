/**
 * @file windows_office_backend.cpp
 * The Windows PowerPoint backend.
 *
 * This is the only file that knows the public API is sitting on top of
 * reverse-engineered Office internals. It adapts the validated native texture
 * store to the platform-independent `Backend` interface, and it converts that
 * store's exceptions into `BackendResult` values so nothing can escape into the
 * C ABI.
 *
 * The safety model is unchanged and lives one level down, in
 * `experiments/exp_internal_blip`: Office build checks, structural PPCORE
 * wrapper validation, OART and GFX vtable checks, per-function signature bytes,
 * and receiver re-resolution on every apply with no cached receiver pointer. A
 * host this backend does not recognise fails closed, and the failure surfaces
 * here as UnsupportedHost or UnsupportedBuild rather than as a crash.
 */

#include "backend.hpp"

#include <windows.h>
#include <oleauto.h>

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include "windows_office/native_texture.hpp"

#include <cstring>
#include <new>

namespace bb {
namespace {

/// Public Office enumeration values, not private ABI offsets.
constexpr long kAutoShape = 1;
constexpr long kFreeform = 5;

/**
 * Maps an internal HRESULT onto the backend vocabulary.
 *
 * The distinction that matters to a caller is "your Office build is not
 * supported" versus "you passed something wrong", so E_NOTIMPL - which the
 * guards use for every unvalidated layout - becomes UnsupportedBuild.
 */
BackendStatus StatusFor(HRESULT hr) {
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
    if (hr == BB_E_TEXTURE_NOT_FOUND) {
        return BackendStatus::InvalidHandle;
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

/**
 * Turns a caller-supplied pointer into a Shape we are willing to fill.
 *
 * VBA hands over `ObjPtr(shape)`, a raw IDispatch with no reference taken, so
 * this validates rather than trusts: the pointer must be readable, must answer
 * QueryInterface for IDispatch, and the object must report a Shape type this
 * backend supports. The reference QueryInterface hands back is released before
 * returning - the caller's own reference is what keeps the Shape alive for the
 * duration of the call.
 */
IDispatch* RequireFillableShape(void* shape) {
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
    // The QueryInterface reference is only needed while the type is checked.
    struct Release {
        IDispatch* value;
        ~Release() { value->Release(); }
    } release{dispatch};

    const long type = get(dispatch, L"Type").integer();
    if (type != kAutoShape && type != kFreeform) {
        throw Error(E_INVALIDARG, "Only AutoShape and Freeform targets are supported");
    }
    return dispatch;
}

/**
 * The Windows implementation.
 *
 * It holds no state of its own: the texture store lives in the native module so
 * that the COM Automation surface and this C ABI address the same textures and
 * cannot disagree about what is loaded.
 */
class WindowsOfficeBackend final : public Backend {
public:
    const char* Name() const noexcept override { return "windows-office-native"; }

    BackendResult Probe() noexcept override {
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return BackendResult::Failure(
                BackendStatus::UnsupportedHost,
                "BlipBridge's accelerated backend runs inside PowerPoint for Windows; "
                "this process is not PowerPoint");
        }
        if (!nativeTextureBackendAvailable()) {
            return BackendResult::Failure(
                BackendStatus::UnsupportedBuild,
                "This Office build has not been validated. The accelerated backend "
                "depends on internal layouts that are checked per build, and refuses "
                "to run against an unrecognised one");
        }
        return BackendResult::Success();
    }

    BackendCapabilities Capabilities() const noexcept override {
        BackendCapabilities capabilities;
        // Every capability here needs a live PowerPoint, including the donor
        // fallback, which drives PickUp/Apply through Automation. Reporting any
        // of them in a process that is not PowerPoint would be a lie a caller
        // could act on.
        if (!GetModuleHandleW(L"POWERPNT.EXE")) {
            return capabilities;
        }
        capabilities.pickUpFallback = true;
        const bool available = nativeTextureBackendAvailable();
        capabilities.nativeBackend = available;
        capabilities.memoryImage = available;
        capabilities.cachedTexture = available;
        capabilities.batchApply = available;
        capabilities.rawPixels = available;
        return capabilities;
    }

    BackendResult LoadTexture(const std::uint8_t* bytes, std::size_t length,
                              std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        if (!bytes || length == 0 || !out) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Image bytes and an output handle are required");
        }
        if (length > 0xFFFFFFFFull) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Image is too large");
        }
        return Guarded([&] {
            // The store copies the bytes, so this array only has to outlive the call.
            SAFEARRAYBOUND bound{static_cast<ULONG>(length), 0};
            SAFEARRAY* array = SafeArrayCreate(VT_UI1, 1, &bound);
            if (!array) {
                throw std::bad_alloc();
            }
            struct Destroy {
                SAFEARRAY* value;
                ~Destroy() { SafeArrayDestroy(value); }
            } destroy{array};

            void* raw = nullptr;
            check(SafeArrayAccessData(array, &raw), "SafeArrayAccessData");
            std::memcpy(raw, bytes, length);
            SafeArrayUnaccessData(array);
            *out = static_cast<std::uint64_t>(nativeTextureLoad(array));
        });
    }

    BackendResult LoadTexturePixels(const std::uint8_t* pixels, std::uint32_t width,
                                    std::uint32_t height, std::int32_t stride,
                                    std::uint64_t* out) noexcept override {
        if (out) {
            *out = 0;
        }
        if (!pixels || width == 0 || height == 0 || !out) {
            return BackendResult::Failure(BackendStatus::InvalidArgument,
                                          "Pixels, dimensions and an output handle are required");
        }
        // A stride smaller than one row of BGRA would read past every row.
        if (stride < static_cast<std::int32_t>(width) * 4) {
            return BackendResult::Failure(
                BackendStatus::InvalidArgument,
                "Stride must be at least width*4 bytes for BGRA32");
        }
        return Guarded([&] {
            *out = static_cast<std::uint64_t>(
                nativeTextureLoadPixels(pixels, width, height, stride));
        });
    }

    BackendResult ApplyTexture(void* shape, std::uint64_t texture) noexcept override {
        return Guarded([&] {
            IDispatch* dispatch = RequireFillableShape(shape);
            // Fill is fetched per call; the receiver behind it is resolved inside
            // and never cached, because a deleted Shape still passes every check.
            nativeTextureApply(get(dispatch, L"Fill").obj(), static_cast<long>(texture));
        });
    }

    BackendResult ReleaseTexture(std::uint64_t texture) noexcept override {
        return Guarded([&] { nativeTextureRelease(static_cast<long>(texture)); });
    }

    void ClearTextures() noexcept override { nativeTextureClear(); }

    std::size_t TextureCount() const noexcept override {
        return static_cast<std::size_t>(nativeTextureCount());
    }
};

} // namespace

std::unique_ptr<Backend> CreateBackend() {
    return std::make_unique<WindowsOfficeBackend>();
}

} // namespace bb
