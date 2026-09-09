#pragma once
#include <blipbridge/dispatch.hpp>
#include <atomic>
#include <map>

namespace bb {
/** Stable Automation IDs. Append methods without renumbering existing entries. */
enum class DispatchId : DISPID {
    GetVersion = 1,
    GetBackendName,
    GetCapabilities,
    RegisterTextureShape,
    ApplyTexture,
    ReleaseTexture,
    ClearTextures,
    GetTextureCount,
    GetLastError,
    LoadTexture,
    SetImageBytes,
    RunBenchmarks,
    GetHostProcessId,
    TraceUserPicture,
    RunFocusedBenchmarks,
    MemoryFillExperiment,
    RunStress,
    TraceCachedApply
};

/**
 * IDTExtensibility2 ABI for the explicitly connected research COM add-in.
 * Method order and 32-bit enum parameters follow Office's public add-in ABI.
 */
struct AddinExtensibility : IDispatch {
    virtual HRESULT STDMETHODCALLTYPE OnConnection(
        IDispatch*, long, IDispatch*, SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(long, SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) = 0;
};

/** Borrows DISPPARAMS for Invoke; extracted Values own dereferenced copies. */
class AutomationArguments {
public:
    explicit AutomationArguments(DISPPARAMS* parameters) : parameters_(parameters) {}
    void RequireCount(UINT expected) const;
    Value At(UINT index) const;
private:
    DISPPARAMS* parameters_;
};

/**
 * STA Automation facade for the explicit whole-style donor fallback.
 * Values retain donor COM wrappers, not independent BLIPs. Clear before closing
 * the donor presentation. Research methods never implement production byte APIs.
 */
class Engine final : public AddinExtensibility {
public:
    Engine();
    ~Engine();
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID interfaceId, void** result) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count) override;
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override;
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        REFIID interfaceId, LPOLESTR* names, UINT count, LCID, DISPID* ids) override;
    HRESULT STDMETHODCALLTYPE Invoke(
        DISPID id, REFIID interfaceId, LCID, WORD flags, DISPPARAMS* parameters,
        VARIANT* result, EXCEPINFO* exception, UINT*) override;
    HRESULT STDMETHODCALLTYPE OnConnection(
        IDispatch*, long, IDispatch* addin, SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnDisconnection(long, SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) override;
private:
    Value Dispatch(DispatchId id, const AutomationArguments& arguments);
    Value DispatchResearch(DispatchId id, const AutomationArguments& arguments);
    /** Retains a normal picture-filled donor; never decodes source bytes. */
    long RegisterTextureShape(Value donor);
    /** STA only. Preserves geometry but transfers the entire donor style. */
    void ApplyTexture(IDispatch* destination, long handle);
    void ReleaseTexture(long handle);
    HRESULT ReportError(const Error& error, EXCEPINFO* exception) noexcept;

    std::atomic<ULONG> references_{1};
    const DWORD owningThread_ = GetCurrentThreadId();
    std::map<long, Value> textures_;
    long nextHandle_ = 1;
    std::wstring lastError_;
};
} // namespace bb
