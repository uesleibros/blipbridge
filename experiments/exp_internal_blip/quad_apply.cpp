/**
 * @file quad_apply.cpp
 * Research driver: warp an image onto a quad and apply it to a Shape.
 *
 * This exists to close the loop between two things that were established
 * separately - that Office maps a picture fill onto the Shape's bounding box
 * (tools/probe_freeform_uv_mapping.ps1), and that src/image/warp.cpp produces a
 * correct projective mapping (tests/warp_contract.cpp) - by doing both and
 * looking at what PowerPoint actually renders.
 *
 * It warps from a **file**, not from a texture handle, and that is not a
 * shortcut. A BlipBridgeTexture holds two opaque GFX pointers and no CPU pixels
 * (see native_texture.cpp), so there is nothing in a handle to warp. Any public
 * quad API has to take its source from encoded bytes, a file or raw BGRA for the
 * same reason, and this driver is shaped like the API it is testing.
 *
 * Research only: the public signature is not frozen, and nothing here is
 * reachable from the C ABI.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/image/pipeline.hpp"
#include "../../src/image/warp.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <sstream>
#include <vector>

namespace {

/// Reads eight doubles - four x,y pairs - out of a SAFEARRAY the harness built.
void ReadQuad(SAFEARRAY* points, bb::image::PointF quad[4]) {
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
        } else if (type == VT_VARIANT) {
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
        } else {
            throw bb::Error(E_INVALIDARG, "Quad coordinates must be numbers");
        }
        (index % 2 == 0 ? quad[index / 2].x : quad[index / 2].y) = static_cast<float>(value);
    }
}

} // namespace

std::wstring warpApplyQuadFromFile(IDispatch* shape, const std::wstring& path, SAFEARRAY* points) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }

    bb::image::PointF quad[4]{};
    ReadQuad(points, quad);

    // Decode once. No crop or resize: the warp is the whole transformation, and
    // resizing first would only throw away detail the warp could have used.
    bb::image::PrepareRequest request;
    const bb::image::PrepareResult decoded = bb::image::PrepareFile(path.c_str(), request);
    if (!decoded.ok()) {
        throw bb::Error(E_INVALIDARG,
                        std::string("Could not read the image: ") +
                            bb::image::DescribeDecodeStatus(decoded.decode));
    }

    const bb::image::WarpResult warped = bb::image::WarpQuad(decoded.image.pixels.data(),
                                                             decoded.image.width,
                                                             decoded.image.height,
                                                             decoded.image.stride(),
                                                             quad,
                                                             0,
                                                             0,
                                                             bb::image::ScaleFilter::Bilinear);
    if (!warped.ok()) {
        throw bb::Error(E_INVALIDARG,
                        std::string("Could not warp onto that quad: ") +
                            bb::image::DescribeWarpStatus(warped.status));
    }

    // A perfectly ordinary texture from here on: the perspective is already in
    // the pixels, so Office's own box mapping is the identity and the Shape's
    // path clips whatever falls outside the quad.
    const bb::office::TextureRef texture =
        bb::office::CreateTextureFromPixels(warped.image.pixels.data(),
                                            warped.image.width,
                                            warped.image.height,
                                            warped.image.stride());
    bb::Value fill = bb::get(shape, L"Fill");
    bb::office::ApplyTextureRef(fill.obj(), texture);

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(2);
    out << L"sourceWidth=" << decoded.image.width << L";sourceHeight=" << decoded.image.height
        << L";rasterWidth=" << warped.image.width << L";rasterHeight=" << warped.image.height
        << L";left=" << warped.left << L";top=" << warped.top << L";width=" << warped.width
        << L";height=" << warped.height << L';';
    return out.str();
}
