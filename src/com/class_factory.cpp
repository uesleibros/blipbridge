#include "server.hpp"

#include <new>
#include <unknwn.h>

namespace bb {
ServerLifetime& GetServerLifetime() noexcept {
    static ServerLifetime lifetime;
    return lifetime;
}

/** Standard nonaggregating COM factory; each live factory prevents DLL unload. */
class ClassFactory final : public IClassFactory {
  public:
    ClassFactory() {
        ++GetServerLifetime().objects;
    }

    ~ClassFactory() {
        --GetServerLifetime().objects;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID interfaceId, void** result) override {
        if (!result) {
            return E_POINTER;
        }
        *result = nullptr;
        if (interfaceId == IID_IUnknown || interfaceId == IID_IClassFactory) {
            *result = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return ++references_;
    }

    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG remaining = --references_;
        if (remaining == 0) {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer,
                                             REFIID interfaceId,
                                             void** result) override {
        if (!result) {
            return E_POINTER;
        }
        *result = nullptr;
        if (outer) {
            return CLASS_E_NOAGGREGATION;
        }
        return CreateEngine(interfaceId, result);
    }

    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock) {
            ++GetServerLifetime().locks;
        } else {
            --GetServerLifetime().locks;
        }
        return S_OK;
    }

  private:
    std::atomic<ULONG> references_{1};
};
} // namespace bb

/** COM activation entry point. Transfers one factory interface reference. */
extern "C" __declspec(dllexport) HRESULT __stdcall
DllGetClassObject(REFCLSID classId, REFIID interfaceId, void** result) {
    if (!result) {
        return E_POINTER;
    }
    *result = nullptr;
    if (classId != bb::kEngineClsid) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    try {
        auto* factory = new bb::ClassFactory;
        const HRESULT status = factory->QueryInterface(interfaceId, result);
        factory->Release();
        return status;
    } catch (const std::bad_alloc&) {
        return E_OUTOFMEMORY;
    } catch (...) {
        return E_UNEXPECTED;
    }
}

/** Unload is permitted only after all Engine/factory references and locks end. */
extern "C" __declspec(dllexport) HRESULT __stdcall DllCanUnloadNow() {
    const auto& lifetime = bb::GetServerLifetime();
    return lifetime.objects == 0 && lifetime.locks == 0 ? S_OK : S_FALSE;
}
