#include "engine.hpp"

#include "../backend/windows_office/native_texture.hpp"
#include "server.hpp"

#include <blipbridge/errors.hpp>
#include <cstring>

namespace bb {
namespace {
constexpr IID kAddinExtensibilityIid = {
    0xb65ad801, 0xabaf, 0x11d0, {0xbb, 0x8b, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x44}};

struct DispatchEntry {
    const wchar_t* name;
    DispatchId id;
};

constexpr DispatchEntry kDispatchEntries[] = {
    {L"GetVersion", DispatchId::GetVersion},
    {L"GetBackendName", DispatchId::GetBackendName},
    {L"GetCapabilities", DispatchId::GetCapabilities},
    {L"RegisterTextureShape", DispatchId::RegisterTextureShape},
    {L"ApplyTexture", DispatchId::ApplyTexture},
    {L"ReleaseTexture", DispatchId::ReleaseTexture},
    {L"ClearTextures", DispatchId::ClearTextures},
    {L"GetTextureCount", DispatchId::GetTextureCount},
    {L"GetLastError", DispatchId::GetLastError},
    {L"LoadTexture", DispatchId::LoadTexture},
    {L"SetImageBytes", DispatchId::SetImageBytes},
    {L"RunBenchmarks", DispatchId::RunBenchmarks},
    {L"GetHostProcessId", DispatchId::GetHostProcessId},
    {L"TraceUserPicture", DispatchId::TraceUserPicture},
    {L"RunFocusedBenchmarks", DispatchId::RunFocusedBenchmarks},
    {L"MemoryFillExperiment", DispatchId::MemoryFillExperiment},
    {L"RunStress", DispatchId::RunStress},
    {L"TraceCachedApply", DispatchId::TraceCachedApply},
    {L"InspectFillReceiver", DispatchId::InspectFillReceiver},
    {L"LoadCachedImageExperiment", DispatchId::LoadCachedImageExperiment},
    {L"NativeApplyExperiment", DispatchId::NativeApplyExperiment},
    {L"NativeApplyReuseExperiment", DispatchId::NativeApplyReuseExperiment},
    {L"InspectTexture", DispatchId::InspectTexture},
    {L"BenchmarkNativeTexture", DispatchId::BenchmarkNativeTexture},
    {L"BenchmarkTextureBatch", DispatchId::BenchmarkTextureBatch},
    {L"BenchmarkApplySkip", DispatchId::BenchmarkApplySkip},
    {L"MeasureApplyCostFactors", DispatchId::MeasureApplyCostFactors},
    {L"ProfileResolveStages", DispatchId::ProfileResolveStages},
    {L"ApplyTextureIfChanged", DispatchId::ApplyTextureIfChanged},
    {L"SplitTransactionApply", DispatchId::SplitTransactionApply},
    {L"ApplyChangeOnly", DispatchId::ApplyChangeOnly},
    {L"ApplyCachedImageToFill", DispatchId::ApplyCachedImageToFill},
    {L"ApplyTextureToRange", DispatchId::ApplyTextureToRange},
    {L"ApplyTextureRange", DispatchId::ApplyTextureRange},
    {L"PixelTextureExperiment", DispatchId::PixelTextureExperiment},
    {L"BenchmarkPixelLoad", DispatchId::BenchmarkPixelLoad},
    {L"ProfileFillStages", DispatchId::ProfileFillStages},
    {L"ProbeShapeCompatibility", DispatchId::ProbeShapeCompatibility},
    {L"ApplyTextureUnrestricted", DispatchId::ApplyTextureUnrestricted},
    {L"ApplyPicture", DispatchId::ApplyPicture},
    {L"InvalidateShape", DispatchId::InvalidateShape},
    {L"ClearPictureCache", DispatchId::ClearPictureCache},
    {L"PictureCacheStats", DispatchId::PictureCacheStats},
    {L"ProbeShapePolicy", DispatchId::ProbeShapePolicy},
    {L"ProfileApplyStages", DispatchId::ProfileApplyStages}};

/**
 * Highest DISPID GetIDsOfNames can hand out. Invoke must accept every one of
 * them, so this is derived from the table rather than written out: a hard-coded
 * bound silently rejects any member appended to the enum.
 */
constexpr DISPID kHighestDispatchId = [] {
    DISPID highest = 0;
    for (const auto& entry : kDispatchEntries) {
        highest = std::max(highest, static_cast<DISPID>(entry.id));
    }
    return highest;
}();
} // namespace

void AutomationArguments::RequireCount(UINT expected) const {
    if (!parameters_ || parameters_->cArgs != expected) {
        throw Error(DISP_E_BADPARAMCOUNT, "Wrong argument count");
    }
}

Value AutomationArguments::At(UINT index) const {
    if (!parameters_ || index >= parameters_->cArgs) {
        throw Error(DISP_E_BADPARAMCOUNT, "Missing argument");
    }
    if (!parameters_->rgvarg) {
        throw Error(E_POINTER, "Missing argument storage");
    }
    Value value;
    // Automation arguments are reversed, and VBA may supply VT_BYREF values.
    check(VariantCopyInd(&value.v, &parameters_->rgvarg[parameters_->cArgs - 1 - index]),
          "Argument");
    return value;
}

Engine::Engine() {
    ++GetServerLifetime().objects;
    OutputDebugStringW(L"BlipBridge 0.1: PickupApplyFallback\n");
}

Engine::~Engine() {
    // Last line of defence for texture lifetime. Office and GFX are still loaded
    // here; an outstanding cached-image reference at process teardown would be a
    // leak, so nothing is left to chance even if OnDisconnection never ran.
    ClearTextures();
    --GetServerLifetime().objects;
}

HRESULT Engine::QueryInterface(REFIID interfaceId, void** result) {
    if (!result) {
        return E_POINTER;
    }
    *result = nullptr;
    if (interfaceId == IID_IUnknown || interfaceId == IID_IDispatch ||
        interfaceId == kAddinExtensibilityIid) {
        *result = static_cast<AddinExtensibility*>(this);
        AddRef();
        return S_OK;
    }
    return E_NOINTERFACE;
}

ULONG Engine::AddRef() {
    return ++references_;
}

ULONG Engine::Release() {
    const ULONG remaining = --references_;
    if (remaining == 0) {
        delete this;
    }
    return remaining;
}

HRESULT Engine::GetTypeInfoCount(UINT* count) {
    if (!count) {
        return E_POINTER;
    }
    *count = 0;
    return S_OK;
}

HRESULT Engine::GetTypeInfo(UINT, LCID, ITypeInfo**) {
    return E_NOTIMPL;
}

HRESULT Engine::GetIDsOfNames(REFIID interfaceId, LPOLESTR* names, UINT count, LCID, DISPID* ids) {
    if (interfaceId != IID_NULL) {
        return DISP_E_UNKNOWNINTERFACE;
    }
    if (!names || !ids) {
        return E_POINTER;
    }
    for (UINT index = 0; index < count; ++index) {
        ids[index] = DISPID_UNKNOWN;
        if (!names[index]) {
            return E_POINTER;
        }
        for (const auto& entry : kDispatchEntries) {
            if (_wcsicmp(names[index], entry.name) == 0) {
                ids[index] = static_cast<DISPID>(entry.id);
                break;
            }
        }
        if (ids[index] == DISPID_UNKNOWN) {
            return DISP_E_UNKNOWNNAME;
        }
    }
    return S_OK;
}

HRESULT Engine::Invoke(DISPID id,
                       REFIID interfaceId,
                       LCID,
                       WORD flags,
                       DISPPARAMS* parameters,
                       VARIANT* result,
                       EXCEPINFO* exception,
                       UINT*) {
    if (interfaceId != IID_NULL) {
        return DISP_E_UNKNOWNINTERFACE;
    }
    if (!(flags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET))) {
        return DISP_E_MEMBERNOTFOUND;
    }
    if (GetCurrentThreadId() != owningThread_) {
        return RPC_E_WRONG_THREAD;
    }
    // Unknown DISPIDs retain the original direct HRESULT, without EXCEPINFO.
    if (id < static_cast<DISPID>(DispatchId::GetVersion) || id > kHighestDispatchId) {
        return DISP_E_MEMBERNOTFOUND;
    }
    try {
        const auto operation = static_cast<DispatchId>(id);
        Value value = Dispatch(operation, AutomationArguments(parameters));
        if (operation != DispatchId::GetLastError) {
            lastError_.clear();
        }
        if (result) {
            VariantInit(result);
            check(VariantCopy(result, &value.v), "Return");
        }
        return S_OK;
    } catch (const Error& error) {
        return ReportError(error, exception);
    } catch (const std::bad_alloc&) {
        return E_OUTOFMEMORY;
    } catch (...) {
        return E_UNEXPECTED;
    }
}

/**
 * Reports what this process can actually do, rather than a compile-time promise.
 *
 * The native backend is guarded to one exact Office build and fails closed
 * everywhere else, so a fixed string would be wrong on most machines. Each flag
 * below is a claim that has been measured; see docs/native_texture.md for the
 * evidence behind each one, and docs/capabilities.md for what they mean.
 */
std::wstring Engine::Capabilities() const {
    const bool nativeBackend = nativeTextureBackendAvailable();
    const wchar_t* native = nativeBackend ? L"True" : L"False";
    std::wstring capabilities;
    // Bytes reach a Shape fill with no temporary image file. Verified against a
    // working control: Fill.UserPicture writes a PNG per call under Content.MSO,
    // the native apply writes none.
    capabilities += L"MemoryImageToFill=";
    capabilities += native;
    // One decoded image applied to many Shapes without re-decoding, proven by
    // creation counts staying equal to the number of LoadTexture calls.
    capabilities += L";CachedTextureApply=";
    capabilities += native;
    capabilities += L";PickUpFallback=True;FillOnly=False;InternalBackend=";
    capabilities += native;
    return capabilities;
}

Value Engine::Dispatch(DispatchId id, const AutomationArguments& arguments) {
    switch (id) {
    case DispatchId::GetVersion:
        arguments.RequireCount(0);
        return Value(L"0.1.0-experimental");
    case DispatchId::GetBackendName:
        arguments.RequireCount(0);
        return Value(L"PickupApplyFallback");
    case DispatchId::GetCapabilities:
        arguments.RequireCount(0);
        return Value(Capabilities().c_str());
    case DispatchId::RegisterTextureShape:
        arguments.RequireCount(1);
        return Value(RegisterTextureShape(arguments.At(0)));
    case DispatchId::ApplyTexture: {
        arguments.RequireCount(2);
        auto destination = arguments.At(0);
        ApplyTexture(destination.obj(), arguments.At(1).integer());
        return Value(-1L);
    }
    case DispatchId::ReleaseTexture:
        arguments.RequireCount(1);
        ReleaseTexture(arguments.At(0).integer());
        return Value();
    case DispatchId::ClearTextures:
        arguments.RequireCount(0);
        ClearTextures();
        return Value();
    case DispatchId::GetTextureCount:
        arguments.RequireCount(0);
        return Value(TextureCount());
    case DispatchId::GetLastError:
        arguments.RequireCount(0);
        return Value(lastError_.c_str());
    case DispatchId::LoadTexture:
        arguments.RequireCount(1);
        return Value(LoadTexture(arguments.At(0)));
    case DispatchId::SetImageBytes:
        throw Error(E_NOTIMPL, kMemoryBackendUnavailableMessage);
    case DispatchId::GetHostProcessId:
        arguments.RequireCount(0);
        return Value(static_cast<long>(GetCurrentProcessId()));
    default:
        return DispatchResearch(id, arguments);
    }
}

HRESULT Engine::ReportError(const Error& error, EXCEPINFO* exception) noexcept {
    // Error reporting itself can allocate. Contain those failures at the COM boundary.
    try {
        lastError_.assign(error.what(), error.what() + std::strlen(error.what()));
        if (exception) {
            *exception = {};
            exception->bstrSource = SysAllocString(L"BlipBridge.Engine");
            exception->bstrDescription = SysAllocString(lastError_.c_str());
            exception->scode = error.hr;
            return DISP_E_EXCEPTION;
        }
        return error.hr;
    } catch (const std::bad_alloc&) {
        return E_OUTOFMEMORY;
    } catch (...) {
        return E_UNEXPECTED;
    }
}

HRESULT Engine::OnConnection(IDispatch*, long, IDispatch* addin, SAFEARRAY**) {
    try {
        put(addin, L"Object", Value(static_cast<IDispatch*>(this)));
        return S_OK;
    } catch (const Error& error) {
        return error.hr;
    } catch (...) {
        return E_UNEXPECTED;
    }
}

HRESULT Engine::OnDisconnection(long, SAFEARRAY**) {
    // Bounds the texture lifetime: every decoded image is released while Office
    // and GFX are still loaded, never left to process teardown.
    ClearTextures();
    return S_OK;
}

HRESULT Engine::OnAddInsUpdate(SAFEARRAY**) {
    return S_OK;
}

HRESULT Engine::OnStartupComplete(SAFEARRAY**) {
    return S_OK;
}

HRESULT Engine::OnBeginShutdown(SAFEARRAY**) {
    textures_.clear();
    return S_OK;
}

HRESULT CreateEngine(REFIID interfaceId, void** result) noexcept {
    try {
        auto* engine = new Engine;
        const HRESULT status = engine->QueryInterface(interfaceId, result);
        engine->Release();
        return status;
    } catch (const std::bad_alloc&) {
        return E_OUTOFMEMORY;
    } catch (...) {
        return E_UNEXPECTED;
    }
}
} // namespace bb
