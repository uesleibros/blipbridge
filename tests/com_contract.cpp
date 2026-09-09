#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include "../src/com/server.hpp"
#include <iostream>
#include <bit>
#include <thread>

namespace {
void Require(bool condition, const char* message) {
    if (!condition) {
        throw std::runtime_error(message);
    }
}

/** Test-owned COM references must die before the loaded server module. */
template<class Interface>
struct OwnedInterface {
    Interface* value = nullptr;
    ~OwnedInterface() {
        if (value) {
            value->Release();
        }
    }
};

struct LoadedModule {
    HMODULE value;
    ~LoadedModule() {
        if (value) {
            FreeLibrary(value);
        }
    }
};

/** Checks both Automation's exception wrapper and the project HRESULT it carries. */
void ExpectAutomationError(IDispatch* engine, DISPID id, HRESULT expected) {
    bb::Value handle(123L);
    DISPPARAMS parameters{&handle.v, nullptr, 1, 0};
    bb::ScopedExceptionInfo exception;
    const HRESULT status = engine->Invoke(
        id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &parameters,
        nullptr, &exception.value, nullptr);
    Require(status == DISP_E_EXCEPTION, "Expected Automation exception");
    Require(exception.value.scode == expected, "Incorrect underlying HRESULT");
}

/**
 * Every name GetIDsOfNames resolves must also be reachable through Invoke.
 * A hard-coded DISPID bound in Invoke once silently rejected a newly appended
 * member while its name still resolved, so the round trip is asserted here.
 * Any HRESULT is acceptable except DISP_E_MEMBERNOTFOUND, which means the ID
 * was refused before dispatch.
 */
void ExpectNameReachableThroughInvoke(IDispatch* engine, const wchar_t* name) {
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    DISPID id = DISPID_UNKNOWN;
    Require(engine->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id) == S_OK,
            "Dispatch name did not resolve");
    DISPPARAMS empty{};
    const HRESULT status = engine->Invoke(
        id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &empty, nullptr, nullptr, nullptr);
    Require(status != DISP_E_MEMBERNOTFOUND, "Resolved name was refused by Invoke");
}
} // namespace

/** Runs without Office: covers COM lifetime, stable IDs, error and STA contracts. */
int wmain(int argc, wchar_t** argv) {
    try {
        Require(argc == 2, "Expected DLL path");
        LoadedModule module{LoadLibraryW(argv[1])};
        Require(module.value != nullptr, "Cannot load BlipBridge");
        using GetClassObject = HRESULT (__stdcall*)(REFCLSID, REFIID, void**);
        using CanUnload = HRESULT (__stdcall*)();
        // These are our documented COM export signatures, not private Office ABIs.
        auto getClass = std::bit_cast<GetClassObject>(
            GetProcAddress(module.value, "DllGetClassObject"));
        auto canUnload = std::bit_cast<CanUnload>(
            GetProcAddress(module.value, "DllCanUnloadNow"));
        Require(getClass && canUnload, "Missing COM exports");
        Require(canUnload() == S_OK, "Fresh server cannot unload");
        {
            OwnedInterface<IClassFactory> factory;
            bb::check(getClass(bb::kEngineClsid, IID_IClassFactory,
                              reinterpret_cast<void**>(&factory.value)), "Get factory");
            Require(canUnload() == S_FALSE, "Factory must retain server");
            Require(factory.value->QueryInterface(IID_IUnknown, nullptr) == E_POINTER,
                    "Factory must reject null output");
            OwnedInterface<IDispatch> engine;
            bb::check(factory.value->CreateInstance(nullptr, IID_IDispatch,
                      reinterpret_cast<void**>(&engine.value)), "Create Engine");
            Require(engine.value->QueryInterface(IID_IUnknown, nullptr) == E_POINTER,
                    "Engine must reject null output");
            Require(bb::call(engine.value, L"GetBackendName").str() == L"PickupApplyFallback",
                    "Backend changed");
            Require(bb::call(engine.value, L"GetTextureCount").integer() == 0,
                    "New Engine must be empty");
            ExpectAutomationError(engine.value, 6, bb::BB_E_TEXTURE_NOT_FOUND);
            // LoadTexture (10) now decodes bytes into a native texture, so an
            // integer argument is a type error rather than an unsupported
            // backend. SetImageBytes (11) is still not implemented.
            ExpectAutomationError(engine.value, 10, E_INVALIDARG);
            ExpectAutomationError(engine.value, 11, E_NOTIMPL);

            // Covers the whole published surface, research members included.
            for (const wchar_t* name : {
                     L"GetVersion", L"GetBackendName", L"GetCapabilities",
                     L"RegisterTextureShape", L"ApplyTexture", L"ReleaseTexture",
                     L"ClearTextures", L"GetTextureCount", L"GetLastError",
                     L"LoadTexture", L"SetImageBytes", L"RunBenchmarks",
                     L"GetHostProcessId", L"TraceUserPicture", L"RunFocusedBenchmarks",
                     L"MemoryFillExperiment", L"RunStress", L"TraceCachedApply",
                     L"InspectFillReceiver", L"LoadCachedImageExperiment",
                     L"NativeApplyExperiment", L"NativeApplyReuseExperiment",
                     L"InspectTexture", L"BenchmarkNativeTexture",
                     L"BenchmarkTextureBatch"}) {
                ExpectNameReachableThroughInvoke(engine.value, name);
            }

            // Intentional invalid raw cross-thread call: tests the early guard only.
            // No Office object is present, and no marshaling or business logic runs.
            HRESULT wrongThread = S_OK;
            std::thread worker([&] {
                DISPPARAMS parameters{};
                wrongThread = engine.value->Invoke(
                    1, IID_NULL, 0, DISPATCH_METHOD, &parameters, nullptr, nullptr, nullptr);
            });
            worker.join();
            Require(wrongThread == RPC_E_WRONG_THREAD, "STA guard failed");
            Require(engine.value->GetIDsOfNames(IID_NULL, nullptr, 1, 0, nullptr) == E_POINTER,
                    "Null name storage accepted");
            factory.value->LockServer(TRUE);
        }
        Require(canUnload() == S_FALSE, "Server lock was lost with factory");
        {
            OwnedInterface<IClassFactory> factory;
            bb::check(getClass(bb::kEngineClsid, IID_IClassFactory,
                              reinterpret_cast<void**>(&factory.value)), "Get factory");
            factory.value->LockServer(FALSE);
        }
        Require(canUnload() == S_OK, "COM references or locks leaked");
        std::cout << "COM contract tests passed\n";
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
