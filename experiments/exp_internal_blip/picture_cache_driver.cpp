/**
 * @file picture_cache_driver.cpp
 * Drives BB_ApplyPicture and its cache controls from the COM research surface.
 *
 * These are thin forwarders to the **real C ABI**, called in process. They exist
 * so `tools/test_picture_cache.ps1` can assert the shipping behaviour rather than
 * a re-implementation of it: everything the harness exercises goes through the
 * same exported functions a VBA caller reaches.
 *
 * Each forwarder turns a non-zero BB_Result into a bb::Error carrying the ABI's
 * own message, so a PowerShell `catch` sees the specific reason - "the image file
 * ..." or "Connectors and lines ..." - and can assert on it.
 *
 * Research only. Nothing in the shipping path calls these.
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

/// The ABI's own explanation of the last failure on this thread.
std::string LastAbiError() {
    const std::uint32_t needed = BB_GetLastError(nullptr, 0);
    if (needed <= 1) {
        return "no message";
    }
    std::vector<char> buffer(needed);
    BB_GetLastError(buffer.data(), needed);
    return std::string(buffer.data());
}

/// Maps an ABI result onto the HRESULT space the Engine reports through EXCEPINFO.
HRESULT HresultFor(BB_Result result) {
    switch (result) {
    case BB_E_INVALID_ARG:
        return E_INVALIDARG;
    case BB_E_UNSUPPORTED_SHAPE:
        return bb::BB_E_SHAPE_CLASS_UNSUPPORTED;
    case BB_E_FILE_NOT_FOUND:
        return bb::BB_E_IMAGE_FILE_MISSING;
    case BB_E_FALLBACK_FAILED:
        return bb::BB_E_FALLBACK_REFUSED;
    case BB_E_INVALID_HANDLE:
        return bb::BB_E_TEXTURE_NOT_FOUND;
    case BB_E_DECODE_FAILED:
        return bb::BB_E_INVALID_IMAGE;
    default:
        return E_FAIL;
    }
}

void RequireOk(BB_Result result, const char* operation) {
    if (result == BB_OK) {
        return;
    }
    std::ostringstream out;
    out << operation << " returned " << result << ": " << LastAbiError();
    throw bb::Error(HresultFor(result), out.str());
}

/// BB_Init is idempotent and re-probes, so every forwarder can simply insist.
void RequireInitialised() {
    RequireOk(BB_Init(), "BB_Init");
}

} // namespace

std::wstring applyPictureThroughAbi(IDispatch* shape, const std::wstring& path) {
    RequireInitialised();
    RequireOk(BB_ApplyPicture(shape, reinterpret_cast<const std::uint16_t*>(path.c_str())),
              "BB_ApplyPicture");
    return L"applied=1;";
}

std::wstring invalidateShapeThroughAbi(IDispatch* shape) {
    RequireInitialised();
    RequireOk(BB_InvalidateShape(shape), "BB_InvalidateShape");
    return L"invalidated=1;";
}

std::wstring clearPictureCacheThroughAbi() {
    RequireInitialised();
    RequireOk(BB_ClearPictureCache(), "BB_ClearPictureCache");
    return L"cleared=1;";
}

std::wstring pictureCacheStatsThroughAbi() {
    RequireInitialised();
    std::uint32_t textures = 0;
    std::uint32_t shapes = 0;
    std::uint64_t skipped = 0;
    RequireOk(BB_GetPictureCacheStats(&textures, &shapes, &skipped), "BB_GetPictureCacheStats");
    std::wostringstream out;
    out << L"textures=" << textures << L";shapes=" << shapes << L";skipped=" << skipped << L';';
    return out.str();
}
