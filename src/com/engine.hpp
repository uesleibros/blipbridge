#pragma once
#include <windows.h>
#include <winnls.h>
#include <atomic>
#include <blipbridge/dispatch.hpp>
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
    TraceCachedApply,
    InspectFillReceiver,
    LoadCachedImageExperiment,
    NativeApplyExperiment,
    NativeApplyReuseExperiment,
    InspectTexture,
    BenchmarkNativeTexture,
    BenchmarkTextureBatch,
    BenchmarkApplySkip,
    MeasureApplyCostFactors,
    ProfileResolveStages,
    ApplyTextureIfChanged,
    SplitTransactionApply,
    ApplyChangeOnly,
    ApplyCachedImageToFill,
    ApplyTextureToRange,
    ApplyTextureRange,
    AbiCapabilities,
    WarpApplyQuad,
    ProbeDynamicTexture,
    LoadTextureScaledAbi,
    ApplyImageQuadAbi,
    ReleaseImageAbi,
    LoadImageAbi,
    PixelTextureExperiment,
    BenchmarkPixelLoad,
    ProfileFillStages,
    ProbeShapeCompatibility,
    ApplyTextureUnrestricted,
    ApplyPicture,
    InvalidateShape,
    ClearPictureCache,
    PictureCacheStats,
    ProbeShapePolicy,
    ProfileApplyStages,
    // Appended for the portable behavioural matrix. Ids are a stable ABI: new
    // ones go on the end, and nothing above this line ever moves.
    LoadTextureBytesAbi,
    LoadTexturePixelsAbi,
    LoadImagePixelsAbi,
    ImageSizeAbi,
    CreateTextureFromImageAbi,
    WarpImageQuadAbi,
    ApplyTextureBatchAbi,
    ClearImagesAbi,
    AbiCounts,
    AbiLifecycle,
    InspectFillStructure,
    InspectModuleIdentities
};

/**
 * IDTExtensibility2 ABI for the explicitly connected research COM add-in.
 * Method order and 32-bit enum parameters follow Office's public add-in ABI.
 */
struct AddinExtensibility : IDispatch {
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*, long, IDispatch*, SAFEARRAY**) = 0;
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
    HRESULT STDMETHODCALLTYPE
    GetIDsOfNames(REFIID interfaceId, LPOLESTR* names, UINT count, LCID, DISPID* ids) override;
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id,
                                     REFIID interfaceId,
                                     LCID,
                                     WORD flags,
                                     DISPPARAMS* parameters,
                                     VARIANT* result,
                                     EXCEPINFO* exception,
                                     UINT*) override;
    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*,
                                           long,
                                           IDispatch* addin,
                                           SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnDisconnection(long, SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) override;
    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) override;

  private:
    Value Dispatch(DispatchId id, const AutomationArguments& arguments);
    Value DispatchResearch(DispatchId id, const AutomationArguments& arguments);
    /** Retains a normal picture-filled donor; never decodes source bytes. */
    long RegisterTextureShape(Value donor);
    /**
     * Decodes bytes once into a reusable native texture. The returned handle
     * owns the decoded image and is independent of any presentation; release it
     * before the Engine goes away. Handles never recycle.
     */
    long LoadTexture(Value bytes);
    /**
     * STA only. A native texture handle applies the decoded image directly and
     * preserves everything but the fill; a donor handle takes the older whole-
     * style transfer path, which preserves geometry but copies the donor style.
     */
    void ApplyTexture(IDispatch* destination, long handle);
    void ReleaseTexture(long handle);
    /** Releases donor references and native textures alike. */
    void ClearTextures() noexcept;
    long TextureCount() const;
    /** Runtime capability string; every flag reflects this host, not a promise. */
    std::wstring Capabilities() const;
    HRESULT ReportError(const Error& error, EXCEPINFO* exception) noexcept;

    std::atomic<ULONG> references_{1};
    const DWORD owningThread_ = GetCurrentThreadId();
    std::map<long, Value> textures_;
    long nextHandle_ = 1;
    std::wstring lastError_;
};
} // namespace bb
