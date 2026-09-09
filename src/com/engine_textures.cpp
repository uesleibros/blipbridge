#include "engine.hpp"
#include "../backend/windows_office/native_texture.hpp"
#include <blipbridge/errors.hpp>
#include <climits>
#include <string>

namespace bb {
namespace {
// Public Office enumeration values, not private ABI offsets.
constexpr long kPictureFill = 6;

/**
 * Rejects Shape classes with no validated picture-fill path.
 *
 * The policy itself is requireFillableShapeClass, which lives with the texture
 * store: both this surface and the C ABI reach the same store, so both must
 * accept exactly the same Shapes. @p message survives only as the caller's own
 * wording for the donor case.
 */
void RequireNormalShape(IDispatch* shape, const char* message) {
    try {
        requireFillableShapeClass(shape);
    } catch (const Error& error) {
        throw Error(error.hr, std::string(message) + ": " + error.what());
    }
}

} // namespace

long Engine::RegisterTextureShape(Value donor) {
    RequireNormalShape(donor.obj(), "Expected AutoShape or Freeform donor");
    if (get(get(donor.obj(), L"Fill").obj(), L"Type").integer() != kPictureFill) {
        throw Error(E_INVALIDARG, "Donor must have picture fill");
    }
    if (nextHandle_ == LONG_MAX) {
        throw Error(E_OUTOFMEMORY, "Handle space exhausted");
    }
    const long handle = nextHandle_++;
    textures_.emplace(handle, std::move(donor));
    return handle;
}

long Engine::LoadTexture(Value bytes) {
    if (bytes.v.vt != (VT_ARRAY | VT_UI1)) {
        throw Error(E_INVALIDARG, "Expected a Byte array");
    }
    return nativeTextureLoad(bytes.v.parray);
}

void Engine::ApplyTexture(IDispatch* destination, long handle) {
    RequireNormalShape(destination, "Target must be AutoShape or Freeform");
    if (nativeTextureOwnsHandle(handle)) {
        // The receiver behind this Fill is resolved inside, per apply, never cached.
        nativeTextureApply(get(destination, L"Fill").obj(), handle);
        return;
    }
    const auto texture = textures_.find(handle);
    if (texture == textures_.end()) {
        throw Error(BB_E_TEXTURE_NOT_FOUND, kTextureNotFoundMessage);
    }
    // Office's picked-up formatting is shared state: refresh it on every call.
    call(texture->second.obj(), L"PickUp");
    call(destination, L"Apply");
}

void Engine::ReleaseTexture(long handle) {
    if (nativeTextureOwnsHandle(handle)) {
        nativeTextureRelease(handle);
        return;
    }
    if (textures_.erase(handle) == 0) {
        throw Error(BB_E_TEXTURE_NOT_FOUND, kTextureNotFoundMessage);
    }
}

void Engine::ClearTextures() noexcept {
    textures_.clear();
    nativeTextureClear();
}

long Engine::TextureCount() const {
    return static_cast<long>(textures_.size()) + nativeTextureCount();
}
} // namespace bb
