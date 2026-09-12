/**
 * @file encode.cpp
 * WIC PNG encoding. See encode.hpp for the format contract and the reasoning.
 */

#include "encode.hpp"

#include "resample.hpp"
#include "wic.hpp"

#include <wincodec.h>

#include <new>

namespace bb::image {
namespace {

/// Maps a WIC failure onto the reason a caller can act on.
EncodeStatus StatusFor(HRESULT hr) noexcept {
    switch (hr) {
    case E_OUTOFMEMORY:
        return EncodeStatus::OutOfMemory;
    case WINCODEC_ERR_COMPONENTNOTFOUND:
        // The PNG encoder is a Windows component, so this means the imaging
        // stack itself is damaged rather than that the request was wrong.
        return EncodeStatus::NoEncoder;
    default:
        return EncodeStatus::EncodeFailed;
    }
}

} // namespace

const char* DescribeEncodeStatus(EncodeStatus status) noexcept {
    switch (status) {
    case EncodeStatus::Ok:
        return "the image was encoded";
    case EncodeStatus::InvalidArgument:
        return "PNG encoding needs pixels, a non-zero width and height, and a stride at least "
               "four bytes per pixel wide";
    case EncodeStatus::NoEncoder:
        return "Windows Imaging Component could not supply a PNG encoder - the imaging stack is "
               "unavailable or damaged";
    case EncodeStatus::UnsupportedSize:
        return "the image has more pixels than this library will encode";
    case EncodeStatus::EncodeFailed:
        return "Windows Imaging Component failed part-way through encoding the image";
    case EncodeStatus::OutOfMemory:
        return "there was not enough memory to encode the image";
    }
    return "the image could not be encoded";
}

EncodeStatus EncodePng(const std::uint8_t* pixels,
                       std::uint32_t width,
                       std::uint32_t height,
                       std::int32_t stride,
                       std::vector<std::uint8_t>& out) noexcept {
    if (!pixels || width == 0 || height == 0) {
        return EncodeStatus::InvalidArgument;
    }
    // A stride narrower than a row would make WIC read into the next row's
    // pixels, or past the buffer entirely on the last one.
    const std::int64_t minimumStride = static_cast<std::int64_t>(width) * 4;
    if (static_cast<std::int64_t>(stride) < minimumStride) {
        return EncodeStatus::InvalidArgument;
    }
    if (static_cast<std::uint64_t>(width) * height > kMaxPixels) {
        return EncodeStatus::UnsupportedSize;
    }

    IWICImagingFactory* factory = Factory();
    if (!factory) {
        return EncodeStatus::NoEncoder;
    }

    // An in-memory stream: the caller decides where the bytes end up, and a
    // failed encode never leaves a half-written file behind for them to find.
    ComPtr<IStream> stream;
    HRESULT hr = CreateStreamOnHGlobal(nullptr, TRUE, stream.put());
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    ComPtr<IWICBitmapEncoder> encoder;
    hr = factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, encoder.put());
    if (FAILED(hr)) {
        return StatusFor(hr);
    }
    hr = encoder->Initialize(stream.get(), WICBitmapEncoderNoCache);
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    ComPtr<IWICBitmapFrameEncode> frame;
    ComPtr<IPropertyBag2> options;
    hr = encoder->CreateNewFrame(frame.put(), options.put());
    if (FAILED(hr)) {
        return StatusFor(hr);
    }
    hr = frame->Initialize(options.get());
    if (FAILED(hr)) {
        return StatusFor(hr);
    }
    hr = frame->SetSize(width, height);
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    /*
     * Ask for BGRA and check what was agreed. WIC negotiates: SetPixelFormat
     * takes the request by reference and writes back the format it will
     * actually use, which for a PNG encoder is BGRA or RGBA depending on the
     * platform's codec. If it comes back as anything but BGRA, handing it our
     * buffer would write an image with red and blue swapped - so the conversion
     * is done by WIC rather than assumed away.
     */
    WICPixelFormatGUID format = GUID_WICPixelFormat32bppBGRA;
    hr = frame->SetPixelFormat(&format);
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    if (format == GUID_WICPixelFormat32bppBGRA) {
        hr = frame->WritePixels(height,
                                static_cast<UINT>(stride),
                                static_cast<UINT>(stride) * height,
                                const_cast<BYTE*>(pixels));
    } else {
        // Wrap the caller's pixels as a bitmap and let WIC convert them into
        // whatever the encoder settled on.
        ComPtr<IWICBitmap> source;
        hr = factory->CreateBitmapFromMemory(width,
                                             height,
                                             GUID_WICPixelFormat32bppBGRA,
                                             static_cast<UINT>(stride),
                                             static_cast<UINT>(stride) * height,
                                             const_cast<BYTE*>(pixels),
                                             source.put());
        if (FAILED(hr)) {
            return StatusFor(hr);
        }
        ComPtr<IWICFormatConverter> converter;
        hr = factory->CreateFormatConverter(converter.put());
        if (FAILED(hr)) {
            return StatusFor(hr);
        }
        hr = converter->Initialize(source.get(),
                                   format,
                                   WICBitmapDitherTypeNone,
                                   nullptr,
                                   0.0,
                                   WICBitmapPaletteTypeCustom);
        if (FAILED(hr)) {
            return StatusFor(hr);
        }
        hr = frame->WriteSource(converter.get(), nullptr);
    }
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    hr = frame->Commit();
    if (FAILED(hr)) {
        return StatusFor(hr);
    }
    hr = encoder->Commit();
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    // Read the finished file back out of the stream.
    HGLOBAL memory = nullptr;
    hr = GetHGlobalFromStream(stream.get(), &memory);
    if (FAILED(hr) || !memory) {
        return StatusFor(FAILED(hr) ? hr : E_FAIL);
    }
    const SIZE_T size = GlobalSize(memory);
    if (size == 0) {
        return EncodeStatus::EncodeFailed;
    }
    const void* data = GlobalLock(memory);
    if (!data) {
        return EncodeStatus::EncodeFailed;
    }
    try {
        out.assign(static_cast<const std::uint8_t*>(data),
                   static_cast<const std::uint8_t*>(data) + size);
    } catch (const std::bad_alloc&) {
        GlobalUnlock(memory);
        return EncodeStatus::OutOfMemory;
    }
    GlobalUnlock(memory);
    return EncodeStatus::Ok;
}

} // namespace bb::image
