#include "engine.hpp"
#include <blipbridge/errors.hpp>
#include <climits>

namespace bb {
namespace {
// Public Office enumeration values, not private ABI offsets.
constexpr long kAutoShape = 1;
constexpr long kFreeform = 5;
constexpr long kPictureFill = 6;

/** Rejects unsupported shape types without changing the document. */
void RequireNormalShape(IDispatch* shape, const char* message) {
    const long type = get(shape, L"Type").integer();
    if (type != kAutoShape && type != kFreeform) {
        throw Error(E_INVALIDARG, message);
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

void Engine::ApplyTexture(IDispatch* destination, long handle) {
    const auto texture = textures_.find(handle);
    if (texture == textures_.end()) {
        throw Error(BB_E_TEXTURE_NOT_FOUND, kTextureNotFoundMessage);
    }
    RequireNormalShape(destination, "Target must be AutoShape or Freeform");
    // Office's picked-up formatting is shared state: refresh it on every call.
    call(texture->second.obj(), L"PickUp");
    call(destination, L"Apply");
}

void Engine::ReleaseTexture(long handle) {
    if (textures_.erase(handle) == 0) {
        throw Error(BB_E_TEXTURE_NOT_FOUND, kTextureNotFoundMessage);
    }
}
} // namespace bb
