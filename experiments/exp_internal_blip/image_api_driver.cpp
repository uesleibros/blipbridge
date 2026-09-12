/**
 * @file image_api_driver.cpp
 * Drives the ABI 5 image surface from the COM research surface.
 *
 * Thin forwarders to the **real C ABI**, called in process, so the harness
 * asserts shipping behaviour rather than a re-implementation of it. Everything
 * the Office suites exercise goes through the same exported functions a VBA
 * caller reaches.
 *
 * Research only in the sense that the *driver* is; the code it calls is the
 * shipped code.
 */

#include "../experiment_api.hpp"

#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <cstdint>
#include <sstream>
#include <string>
#include <vector>

namespace {

std::string LastAbiError() {
    const std::uint32_t needed = BB_GetLastError(nullptr, 0);
    if (needed <= 1) {
        return "no message";
    }
    std::vector<char> buffer(needed);
    BB_GetLastError(buffer.data(), needed);
    return std::string(buffer.data());
}

void RequireOk(BB_Result result, const char* operation) {
    if (result == BB_OK) {
        return;
    }
    std::ostringstream out;
    out << operation << " returned " << result << ": " << LastAbiError();
    throw bb::Error(result == BB_E_UNSUPPORTED_SHAPE ? bb::BB_E_SHAPE_CLASS_UNSUPPORTED : E_FAIL,
                    out.str());
}

void RequireInitialised() {
    RequireOk(BB_Init(), "BB_Init");
}

/// Parses "cropX,cropY,cropW,cropH,transform,targetW,targetH,filter" from a harness.
BB_ImageRequest ReadRequest(const std::wstring& text) {
    BB_ImageRequest request{};
    if (text.empty()) {
        return request;
    }
    std::wistringstream in(text);
    std::uint32_t* fields[] = {&request.cropX,
                               &request.cropY,
                               &request.cropWidth,
                               &request.cropHeight,
                               &request.transform,
                               &request.targetWidth,
                               &request.targetHeight,
                               &request.filter};
    for (std::uint32_t* field : fields) {
        std::wstring token;
        if (!std::getline(in, token, L',')) {
            break;
        }
        *field = static_cast<std::uint32_t>(std::stoul(token));
    }
    return request;
}

/// Reads eight numbers - four x,y pairs - out of a SAFEARRAY the harness built.
void ReadQuad(SAFEARRAY* points, BB_PointF quad[4]) {
    if (!points || SafeArrayGetDim(points) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional array of eight numbers");
    }
    LONG lower = 0;
    LONG upper = 0;
    bb::check(SafeArrayGetLBound(points, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(points, 1, &upper), "SafeArrayGetUBound");
    if (upper - lower + 1 != 8) {
        throw bb::Error(E_INVALIDARG, "A quad needs exactly eight numbers: x0,y0 ... x3,y3");
    }

    VARTYPE type = VT_EMPTY;
    bb::check(SafeArrayGetVartype(points, &type), "SafeArrayGetVartype");
    for (LONG index = 0; index < 8; ++index) {
        LONG at = lower + index;
        double value = 0.0;
        if (type == VT_R8) {
            bb::check(SafeArrayGetElement(points, &at, &value), "SafeArrayGetElement");
        } else {
            VARIANT element{};
            VariantInit(&element);
            bb::check(SafeArrayGetElement(points, &at, &element), "SafeArrayGetElement");
            VARIANT asDouble{};
            VariantInit(&asDouble);
            const HRESULT changed = VariantChangeType(&asDouble, &element, 0, VT_R8);
            VariantClear(&element);
            if (FAILED(changed)) {
                throw bb::Error(E_INVALIDARG, "A quad coordinate was not a number");
            }
            value = asDouble.dblVal;
        }
        (index % 2 == 0 ? quad[index / 2].x : quad[index / 2].y) = static_cast<float>(value);
    }
}

} // namespace

std::wstring loadImageThroughAbi(const std::wstring& path, const std::wstring& request) {
    RequireInitialised();
    const BB_ImageRequest parsed = ReadRequest(request);
    BB_Image image = 0;
    RequireOk(BB_LoadImageFromFile(reinterpret_cast<const std::uint16_t*>(path.c_str()),
                                   request.empty() ? nullptr : &parsed,
                                   &image),
              "BB_LoadImageFromFile");

    std::uint32_t width = 0;
    std::uint32_t height = 0;
    RequireOk(BB_GetImageSize(image, &width, &height), "BB_GetImageSize");

    std::wostringstream out;
    out << L"image=" << image << L";width=" << width << L";height=" << height << L";count="
        << BB_GetImageCount() << L';';
    return out.str();
}

std::wstring releaseImageThroughAbi(long long image) {
    RequireInitialised();
    RequireOk(BB_ReleaseImage(static_cast<BB_Image>(image)), "BB_ReleaseImage");
    std::wostringstream out;
    out << L"released=1;count=" << BB_GetImageCount() << L';';
    return out.str();
}

std::wstring applyImageQuadThroughAbi(IDispatch* shape, long long image, SAFEARRAY* points) {
    RequireInitialised();
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    BB_PointF quad[4]{};
    ReadQuad(points, quad);
    RequireOk(BB_ApplyImageQuad(shape, static_cast<BB_Image>(image), quad, BB_SCALE_BILINEAR),
              "BB_ApplyImageQuad");
    std::wostringstream out;
    // The convenience wrapper owns the texture it made, so the count must be
    // unchanged by the call - that is the leak this asserts against.
    out << L"applied=1;textures=" << BB_GetTextureCount() << L';';
    return out.str();
}

std::wstring loadTextureScaledThroughAbi(const std::wstring& path, const std::wstring& request) {
    RequireInitialised();
    const BB_ImageRequest parsed = ReadRequest(request);
    BB_Handle texture = 0;
    RequireOk(BB_LoadTextureFromFileEx(reinterpret_cast<const std::uint16_t*>(path.c_str()),
                                       request.empty() ? nullptr : &parsed,
                                       &texture),
              "BB_LoadTextureFromFileEx");
    std::wostringstream out;
    // No CPU image is created by this path, which is the whole reason it exists.
    out << L"texture=" << texture << L";images=" << BB_GetImageCount() << L';';
    return out.str();
}

namespace {

/// Borrows a VT_UI1 SAFEARRAY's bytes for the duration of a call.
class BorrowedBytes {
  public:
    explicit BorrowedBytes(SAFEARRAY* array) : array_(array) {
        if (!array_) {
            throw bb::Error(E_INVALIDARG, "Expected a Byte array");
        }
        bb::check(SafeArrayAccessData(array_, &data_), "SafeArrayAccessData");
        long lower = 0;
        long upper = 0;
        SafeArrayGetLBound(array_, 1, &lower);
        SafeArrayGetUBound(array_, 1, &upper);
        length_ = static_cast<std::size_t>(upper - lower + 1);
    }

    ~BorrowedBytes() {
        SafeArrayUnaccessData(array_);
    }

    BorrowedBytes(const BorrowedBytes&) = delete;
    BorrowedBytes& operator=(const BorrowedBytes&) = delete;

    const std::uint8_t* data() const noexcept {
        return static_cast<const std::uint8_t*>(data_);
    }
    std::size_t size() const noexcept {
        return length_;
    }

  private:
    SAFEARRAY* array_ = nullptr;
    void* data_ = nullptr;
    std::size_t length_ = 0;
};

/// The two counts, appended to every reply so a harness can assert deltas.
void AppendCounts(std::wostringstream& out) {
    out << L"textures=" << BB_GetTextureCount() << L";images=" << BB_GetImageCount() << L';';
}

} // namespace

std::wstring loadTexturePixelsThroughAbi(SAFEARRAY* pixels, long width, long height, long stride) {
    RequireInitialised();
    const BorrowedBytes bytes(pixels);
    BB_Handle texture = 0;
    RequireOk(BB_LoadTexturePixels(bytes.data(),
                                   static_cast<std::uint32_t>(width),
                                   static_cast<std::uint32_t>(height),
                                   static_cast<std::int32_t>(stride),
                                   &texture),
              "BB_LoadTexturePixels");
    std::wostringstream out;
    out << L"texture=" << texture << L';';
    AppendCounts(out);
    return out.str();
}

std::wstring loadImagePixelsThroughAbi(SAFEARRAY* pixels, long width, long height, long stride) {
    RequireInitialised();
    const BorrowedBytes bytes(pixels);
    BB_Image image = 0;
    RequireOk(BB_LoadImagePixels(bytes.data(),
                                 static_cast<std::uint32_t>(width),
                                 static_cast<std::uint32_t>(height),
                                 static_cast<std::int32_t>(stride),
                                 nullptr,
                                 &image),
              "BB_LoadImagePixels");
    std::uint32_t gotWidth = 0;
    std::uint32_t gotHeight = 0;
    RequireOk(BB_GetImageSize(image, &gotWidth, &gotHeight), "BB_GetImageSize");
    std::wostringstream out;
    out << L"image=" << image << L";width=" << gotWidth << L";height=" << gotHeight << L';';
    AppendCounts(out);
    return out.str();
}

std::wstring loadTextureBytesThroughAbi(SAFEARRAY* encoded) {
    RequireInitialised();
    const BorrowedBytes bytes(encoded);
    BB_Handle texture = 0;
    RequireOk(BB_LoadTexture(bytes.data(), static_cast<std::uint32_t>(bytes.size()), &texture),
              "BB_LoadTexture");
    std::wostringstream out;
    out << L"texture=" << texture << L';';
    AppendCounts(out);
    return out.str();
}

std::wstring imageSizeThroughAbi(long long image) {
    RequireInitialised();
    std::uint32_t width = 0;
    std::uint32_t height = 0;
    RequireOk(BB_GetImageSize(static_cast<BB_Image>(image), &width, &height), "BB_GetImageSize");
    std::wostringstream out;
    out << L"width=" << width << L";height=" << height << L';';
    return out.str();
}

std::wstring createTextureFromImageThroughAbi(long long image) {
    RequireInitialised();
    BB_Handle texture = 0;
    RequireOk(BB_CreateTextureFromImage(static_cast<BB_Image>(image), &texture),
              "BB_CreateTextureFromImage");
    std::wostringstream out;
    out << L"texture=" << texture << L';';
    AppendCounts(out);
    return out.str();
}

std::wstring warpImageQuadThroughAbi(long long image, SAFEARRAY* points, long filter) {
    RequireInitialised();
    BB_PointF quad[4]{};
    ReadQuad(points, quad);
    BB_Handle texture = 0;
    RequireOk(BB_WarpImageQuad(static_cast<BB_Image>(image),
                               quad,
                               static_cast<std::uint32_t>(filter),
                               &texture),
              "BB_WarpImageQuad");
    std::wostringstream out;
    // The primitive hands the texture to the caller, unlike ApplyImageQuad which
    // owns the one it makes - so here the texture count is expected to rise.
    out << L"texture=" << texture << L';';
    AppendCounts(out);
    return out.str();
}

std::wstring applyTextureBatchThroughAbi(SAFEARRAY* shapes, long long texture) {
    RequireInitialised();
    if (!shapes) {
        throw bb::Error(E_INVALIDARG, "Expected an array of Shapes");
    }
    long lower = 0;
    long upper = 0;
    bb::check(SafeArrayGetLBound(shapes, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(shapes, 1, &upper), "SafeArrayGetUBound");
    const long count = upper - lower + 1;
    if (count <= 0) {
        throw bb::Error(E_INVALIDARG, "The Shape array is empty");
    }

    std::vector<void*> pointers;
    std::vector<BB_Handle> handles;
    pointers.reserve(static_cast<std::size_t>(count));
    handles.reserve(static_cast<std::size_t>(count));
    for (long index = lower; index <= upper; ++index) {
        VARIANT element{};
        VariantInit(&element);
        bb::check(SafeArrayGetElement(shapes, &index, &element), "SafeArrayGetElement");
        if (element.vt != VT_DISPATCH || !element.pdispVal) {
            VariantClear(&element);
            throw bb::Error(E_INVALIDARG, "A batch member is not a Shape");
        }
        pointers.push_back(element.pdispVal);
        handles.push_back(static_cast<BB_Handle>(texture));
        // The VARIANT's reference is released here; the caller's collection is
        // what keeps every Shape alive across the call.
        element.pdispVal->Release();
        element.vt = VT_EMPTY;
    }

    std::uint32_t applied = 0;
    RequireOk(BB_ApplyTextureBatch(pointers.data(),
                                   handles.data(),
                                   static_cast<std::uint32_t>(count),
                                   &applied),
              "BB_ApplyTextureBatch");
    std::wostringstream out;
    out << L"applied=" << applied << L";requested=" << count << L';';
    return out.str();
}

std::wstring clearImagesThroughAbi() {
    RequireInitialised();
    RequireOk(BB_ClearImages(), "BB_ClearImages");
    std::wostringstream out;
    out << L"cleared=1;";
    AppendCounts(out);
    return out.str();
}

std::wstring abiCountsThroughAbi() {
    RequireInitialised();
    std::wostringstream out;
    AppendCounts(out);
    return out.str();
}

std::wstring lifecycleThroughAbi(const std::wstring& action) {
    std::wostringstream out;
    if (action == L"shutdown") {
        RequireOk(BB_Shutdown(), "BB_Shutdown");
        out << L"shutdown=1;";
    } else if (action == L"init") {
        RequireOk(BB_Init(), "BB_Init");
        out << L"init=1;";
        AppendCounts(out);
    } else {
        throw bb::Error(E_INVALIDARG, "Expected 'init' or 'shutdown'");
    }
    return out.str();
}
