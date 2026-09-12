/**
 * @file decode.cpp
 * WIC decoding to canonical BGRA32. See decode.hpp for the format contract.
 */

#include "decode.hpp"

#include "resample.hpp"
#include "wic.hpp"

#include <wincodec.h>

#include <cstdio>
#include <new>

namespace bb::image {
namespace {

/// Maps a WIC failure onto the reason a caller can act on.
DecodeStatus StatusFor(HRESULT hr) noexcept {
    switch (hr) {
    case WINCODEC_ERR_COMPONENTNOTFOUND:
    case WINCODEC_ERR_UNKNOWNIMAGEFORMAT:
    case WINCODEC_ERR_BADHEADER:
        return DecodeStatus::UnknownFormat;
    case WINCODEC_ERR_BADIMAGE:
    case WINCODEC_ERR_BADSTREAMDATA:
    case WINCODEC_ERR_STREAMREAD:
    case WINCODEC_ERR_UNEXPECTEDMETADATATYPE:
        return DecodeStatus::Corrupt;
    case E_OUTOFMEMORY:
        return DecodeStatus::OutOfMemory;
    default:
        // An unrecognised failure from a recognised decoder is damaged data far
        // more often than anything else, and saying "corrupt" is more use than
        // saying "0x88982F60".
        return DecodeStatus::Corrupt;
    }
}

/**
 * Decodes the first frame of @p decoder into @p out as BGRA32.
 *
 * Split out because the byte and file entry points differ only in how they get
 * a decoder, and the part that must be careful is all after that.
 */
DecodeStatus DecodeFirstFrame(IWICBitmapDecoder* decoder, DecodedImage& out) noexcept {
    ComPtr<IWICBitmapFrameDecode> frame;
    HRESULT hr = decoder->GetFrame(0, frame.put());
    if (FAILED(hr) || !frame) {
        return StatusFor(hr);
    }

    /*
     * Convert rather than inspect. WIC will take any source format to BGRA in
     * one step, including palettes, 16-bit channels and CMYK, so asking for the
     * canonical format is both simpler and more correct than a switch over
     * formats this file would have to keep up to date.
     *
     * GUID_WICPixelFormat32bppBGRA is straight alpha. The premultiplied variant
     * exists and is deliberately not used: resample.hpp premultiplies where it
     * needs to, and doing it twice darkens every transparent edge.
     */
    ComPtr<IWICFormatConverter> converter;
    IWICImagingFactory* factory = Factory();
    if (!factory) {
        return DecodeStatus::NoDecoder;
    }
    hr = factory->CreateFormatConverter(converter.put());
    if (FAILED(hr) || !converter) {
        return DecodeStatus::ConversionFailed;
    }
    hr = converter->Initialize(frame.get(),
                               GUID_WICPixelFormat32bppBGRA,
                               WICBitmapDitherTypeNone,
                               nullptr,
                               0.0,
                               WICBitmapPaletteTypeCustom);
    if (FAILED(hr)) {
        return DecodeStatus::ConversionFailed;
    }

    UINT width = 0;
    UINT height = 0;
    hr = converter->GetSize(&width, &height);
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    // Every size that reaches an allocation is checked in 64-bit first. A
    // decoder reporting an absurd size is the one thing between here and a
    // short buffer with a long write.
    const std::uint64_t pixels = static_cast<std::uint64_t>(width) * height;
    if (width == 0 || height == 0 || pixels > kMaxPixels) {
        return DecodeStatus::UnsupportedSize;
    }
    const std::uint64_t bytes = pixels * 4ull;
    const std::uint64_t stride = static_cast<std::uint64_t>(width) * 4ull;
    if (stride > 0x7FFFFFFFull || bytes > 0x7FFFFFFFull) {
        return DecodeStatus::UnsupportedSize;
    }

    DecodedImage decoded;
    decoded.width = width;
    decoded.height = height;
    try {
        decoded.pixels.resize(static_cast<std::size_t>(bytes));
    } catch (const std::bad_alloc&) {
        return DecodeStatus::OutOfMemory;
    }

    hr = converter->CopyPixels(
        nullptr, static_cast<UINT>(stride), static_cast<UINT>(bytes), decoded.pixels.data());
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    out = std::move(decoded);
    return DecodeStatus::Ok;
}

} // namespace

const char* DescribeDecodeStatus(DecodeStatus status) noexcept {
    switch (status) {
    case DecodeStatus::Ok:
        return "The image was decoded";
    case DecodeStatus::NoData:
        return "No image data was given";
    case DecodeStatus::NoDecoder:
        return "Windows Imaging Component is not available in this process";
    case DecodeStatus::UnknownFormat:
        return "The data is not an image format Windows can decode";
    case DecodeStatus::Corrupt:
        return "The image data is damaged or ends before the image does";
    case DecodeStatus::UnsupportedSize:
        return "The image has no pixels, or more than 64 megapixels";
    case DecodeStatus::ConversionFailed:
        return "The image could not be converted to BGRA32";
    case DecodeStatus::OutOfMemory:
        return "Not enough memory to hold the decoded image";
    }
    return "Unknown decode status";
}

DecodeStatus Decode(const std::uint8_t* bytes, std::size_t length, DecodedImage& out) noexcept {
    if (!bytes || length == 0) {
        return DecodeStatus::NoData;
    }
    if (length > kMaxEncodedBytes) {
        return DecodeStatus::UnsupportedSize;
    }

    IWICImagingFactory* factory = Factory();
    if (!factory) {
        return DecodeStatus::NoDecoder;
    }

    // A stream over the caller's buffer: WIC reads it in place and never takes
    // ownership, so the bytes only have to outlive this call.
    ComPtr<IWICStream> stream;
    HRESULT hr = factory->CreateStream(stream.put());
    if (FAILED(hr) || !stream) {
        return DecodeStatus::NoDecoder;
    }
    hr = stream->InitializeFromMemory(const_cast<BYTE*>(bytes), static_cast<DWORD>(length));
    if (FAILED(hr)) {
        return StatusFor(hr);
    }

    ComPtr<IWICBitmapDecoder> decoder;
    hr = factory->CreateDecoderFromStream(
        stream.get(), nullptr, WICDecodeMetadataCacheOnDemand, decoder.put());
    if (FAILED(hr) || !decoder) {
        return StatusFor(hr);
    }
    return DecodeFirstFrame(decoder.get(), out);
}

DecodeStatus DecodeFile(const wchar_t* path, DecodedImage& out) noexcept {
    if (!path || !*path) {
        return DecodeStatus::NoData;
    }

    IWICImagingFactory* factory = Factory();
    if (!factory) {
        return DecodeStatus::NoDecoder;
    }

    /*
     * WIC opens the file itself rather than this reading it into memory first.
     * That keeps a large file from being copied once to be parsed once, and it
     * is WIC's own path for what a decoder needs - but it also means the
     * kMaxEncodedBytes ceiling has to be applied here, before the decoder is
     * handed a path, rather than falling out of a read.
     */
    WIN32_FILE_ATTRIBUTE_DATA attributes{};
    if (!GetFileAttributesExW(path, GetFileExInfoStandard, &attributes)) {
        return DecodeStatus::NoData;
    }
    const std::uint64_t size =
        (static_cast<std::uint64_t>(attributes.nFileSizeHigh) << 32) | attributes.nFileSizeLow;
    if (size == 0) {
        return DecodeStatus::NoData;
    }
    if (size > kMaxEncodedBytes) {
        return DecodeStatus::UnsupportedSize;
    }

    ComPtr<IWICBitmapDecoder> decoder;
    const HRESULT hr = factory->CreateDecoderFromFilename(
        path, nullptr, GENERIC_READ, WICDecodeMetadataCacheOnDemand, decoder.put());
    if (FAILED(hr) || !decoder) {
        // A path that does not exist reports as a missing file rather than as an
        // unknown format, which is what a caller needs to hear.
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ||
            hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ||
            hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED)) {
            return DecodeStatus::NoData;
        }
        return StatusFor(hr);
    }
    return DecodeFirstFrame(decoder.get(), out);
}

} // namespace bb::image
