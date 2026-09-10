#include <blipbridge/errors.hpp>
#include <blipbridge/image_validation.hpp>
#include <cstdint>
#include <cstring>
#include <objbase.h>
#include <vector>
#include <wincodec.h>

namespace bb {
template <class T>
struct Com {
    T* p = nullptr;

    ~Com() {
        if (p) {
            p->Release();
        }
    }

    T** out() {
        return &p;
    }
};

HRESULT validateImage(const BYTE* data, size_t size) {
    constexpr HRESULT invalid = BB_E_INVALID_IMAGE;
    const BYTE png[] = {137, 80, 78, 71, 13, 10, 26, 10};
    bool isPng = size >= 8 && !memcmp(data, png, 8);
    bool isJpeg = size >= 3 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff;
    if (!isPng && !isJpeg) {
        return invalid;
    }
    try {
        Com<IWICImagingFactory> factory;
        auto hr = CoCreateInstance(CLSID_WICImagingFactory,
                                   nullptr,
                                   CLSCTX_INPROC_SERVER,
                                   IID_IWICImagingFactory,
                                   (void**)factory.out());
        if (FAILED(hr)) {
            return hr;
        }
        Com<IWICStream> stream;
        hr = factory.p->CreateStream(stream.out());
        if (FAILED(hr)) {
            return hr;
        }
        hr = stream.p->InitializeFromMemory(const_cast<BYTE*>(data), (DWORD)size);
        if (FAILED(hr)) {
            return invalid;
        }
        Com<IWICBitmapDecoder> decoder;
        hr = factory.p->CreateDecoderFromStream(
            stream.p, nullptr, WICDecodeMetadataCacheOnLoad, decoder.out());
        if (FAILED(hr)) {
            return invalid;
        }
        Com<IWICBitmapFrameDecode> frame;
        hr = decoder.p->GetFrame(0, frame.out());
        if (FAILED(hr)) {
            return invalid;
        }
        UINT width = 0, height = 0;
        hr = frame.p->GetSize(&width, &height);
        if (FAILED(hr) || !width || !height || uint64_t(width) * height > 16 * 1024 * 1024) {
            return invalid;
        }
        Com<IWICFormatConverter> converted;
        hr = factory.p->CreateFormatConverter(converted.out());
        if (FAILED(hr)) {
            return hr;
        }
        hr = converted.p->Initialize(frame.p,
                                     GUID_WICPixelFormat32bppBGRA,
                                     WICBitmapDitherTypeNone,
                                     nullptr,
                                     0,
                                     WICBitmapPaletteTypeCustom);
        if (FAILED(hr)) {
            return invalid;
        }
        std::vector<BYTE> row(size_t(width) * 4);
        for (UINT y = 0; y < height; y++) {
            WICRect area{0, (INT)y, (INT)width, 1};
            hr = converted.p->CopyPixels(&area, width * 4, (UINT)row.size(), row.data());
            if (FAILED(hr)) {
                return invalid;
            }
        }
        return S_OK;
    } catch (const std::bad_alloc&) {
        return E_OUTOFMEMORY;
    } catch (...) {
        return E_UNEXPECTED;
    }
}
} // namespace bb
