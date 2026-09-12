#pragma once
/**
 * @file wic.hpp
 * The Windows Imaging Component plumbing shared by decoding and encoding.
 *
 * Two things live here, and they are here rather than in either .cpp because
 * both files need them and a second copy of the factory would be a second
 * factory - one more `CoCreateInstance` per thread, for nothing.
 *
 * Internal to `src/image`. Nothing in the public headers refers to it.
 */

// MinGW needs the Windows base types before anything COM.
#include <windows.h>

#include <wincodec.h>

namespace bb::image {

/// Releases a COM pointer on scope exit, so every early return stays correct.
template <typename T>
class ComPtr {
  public:
    ComPtr() = default;
    ComPtr(const ComPtr&) = delete;
    ComPtr& operator=(const ComPtr&) = delete;

    ~ComPtr() {
        if (value_) {
            value_->Release();
        }
    }

    T** put() noexcept {
        return &value_;
    }

    T* get() const noexcept {
        return value_;
    }

    T* operator->() const noexcept {
        return value_;
    }

    explicit operator bool() const noexcept {
        return value_ != nullptr;
    }

  private:
    T* value_ = nullptr;
};

/**
 * The imaging factory, created once per thread that asks for one.
 *
 * WIC objects are apartment-bound and this library is single-threaded-apartment
 * throughout, so a thread-local factory is both correct and one fewer
 * CoCreateInstance per image. It is deliberately never released: it lives as
 * long as the thread, which is what a factory is for, and releasing it during
 * DLL teardown would mean touching COM at a moment when COM may be gone.
 */
inline IWICImagingFactory* Factory() noexcept {
    static thread_local IWICImagingFactory* factory = nullptr;
    if (factory) {
        return factory;
    }
    // Succeeds whether or not this thread has already called CoInitialize: WIC
    // is a free-threaded-marshalled in-proc server, so an uninitialised
    // apartment is the one case worth retrying after initialising.
    HRESULT hr = CoCreateInstance(
        CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
    if (hr == CO_E_NOTINITIALIZED) {
        if (SUCCEEDED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
            hr = CoCreateInstance(
                CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
        }
    }
    return SUCCEEDED(hr) ? factory : nullptr;
}

} // namespace bb::image
