/**
 * @file engine_textures_portable.cpp
 * The research COM surface's texture methods, over the portable backend.
 *
 * Compiled instead of `engine_textures.cpp` wherever the accelerated backend is
 * not. The two files implement the same five `Engine` methods and must behave
 * the same way, because the Office harnesses drive this surface and their
 * assertions are about the library's contract, not about which backend is
 * underneath.
 *
 * There is one deliberate difference in how they get there. The accelerated
 * version calls the native texture store directly, because that store is what it
 * is instrumenting. This one goes through the **public C ABI** - `BB_LoadTexture`,
 * `BB_ApplyTexture` and so on - which is a strictly better arrangement for the
 * job this file has: it means a harness driving the COM surface is exercising
 * the exact entry points a caller uses, so the two cannot disagree about
 * anything, including their error text.
 *
 * The donor `PickUp`/`Apply` path is unchanged and lives here in full. It never
 * needed a backend - it is Office formatting one Shape from another - so it is
 * the same code on both.
 */

#include "../backend/portable_office/portable_texture.hpp"
#include "../backend/windows_office/shape_policy.hpp"
#include "engine.hpp"

#include <blipbridge/blipbridge.h>
#include <blipbridge/errors.hpp>

#include <climits>
#include <string>
#include <vector>

namespace bb {
namespace {
// Public Office enumeration values, not private ABI offsets.
constexpr long kPictureFill = 6;

/**
 * Rejects Shape classes with no picture-fill path at all.
 *
 * The policy itself is bb::office::RequireNativePictureFillTarget - the same
 * authority the C ABI asks - so this surface and that one accept exactly the
 * same Shapes. @p message survives only as the caller's own wording.
 */
void RequireNormalShape(IDispatch* shape, const char* message) {
    try {
        office::RequireNativePictureFillTarget(shape);
    } catch (const Error& error) {
        throw Error(error.hr, std::string(message) + ": " + error.what());
    }
}

/// Turns a C ABI failure into the exception this surface reports, with its text.
void RequireOk(BB_Result result, const char* operation) {
    if (result == BB_OK) {
        return;
    }
    const std::uint32_t needed = BB_GetLastError(nullptr, 0);
    std::string message(operation);
    if (needed > 1) {
        std::vector<char> buffer(needed);
        BB_GetLastError(buffer.data(), needed);
        message += ": ";
        message += buffer.data();
    }
    // BB_E_UNSUPPORTED_SHAPE is the one code a harness distinguishes, because a
    // refused Shape class is a result rather than a fault.
    throw Error(result == BB_E_UNSUPPORTED_SHAPE ? BB_E_SHAPE_CLASS_UNSUPPORTED : E_FAIL,
                message);
}

/**
 * Makes sure the library is initialised before the first ABI call.
 *
 * The accelerated version of this file never needed it: it drove the native
 * store directly, and that store initialises itself. Going through the C ABI
 * means playing by the C ABI's rules, and the first of them is BB_Init. It is
 * documented as safe to call repeatedly, so this asks every time rather than
 * keeping a flag that could disagree with the library's own.
 */
void RequireInitialised() {
    RequireOk(BB_Init(), "BB_Init");
}

/*
 * Where the initialise call goes, and why it is not simply first.
 *
 * BB_Init fails outside PowerPoint, by design. If it ran before the argument
 * checks, then off-host every wrong argument would report "not running in
 * PowerPoint" instead of what was actually wrong - and tests/com_contract.cpp
 * runs off-host precisely to pin those errors down. So each method below
 * validates what it was given first and initialises only when it is about to
 * enter the C ABI.
 */

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
    RequireInitialised();
    SAFEARRAY* array = bytes.v.parray;
    void* raw = nullptr;
    check(SafeArrayAccessData(array, &raw), "SafeArrayAccessData");
    long lower = 0;
    long upper = 0;
    SafeArrayGetLBound(array, 1, &lower);
    SafeArrayGetUBound(array, 1, &upper);
    const std::size_t length = static_cast<std::size_t>(upper - lower + 1);

    BB_Handle handle = 0;
    const BB_Result result = BB_LoadTexture(
        static_cast<const std::uint8_t*>(raw), static_cast<std::uint32_t>(length), &handle);
    SafeArrayUnaccessData(array);
    RequireOk(result, "BB_LoadTexture");

    // The research surface's handles are `long`, which is what its Automation
    // signature has always been. Texture handles start at 0x1000000 and rise one
    // at a time, so this cannot truncate in any run that could also finish.
    return static_cast<long>(handle);
}

void Engine::ApplyTexture(IDispatch* destination, long handle) {
    RequireNormalShape(destination, "Target must be AutoShape or Freeform");
    // Routed on the handle *space*, not on ownership: a value from the texture
    // space that this store does not own is a stale or invented texture handle,
    // and the store says which. Falling through to the donor map instead would
    // answer "not found" and lose that.
    if (portable::IsTextureHandle(static_cast<std::uint64_t>(handle))) {
        RequireInitialised();
        RequireOk(BB_ApplyTexture(destination, static_cast<BB_Handle>(handle)), "BB_ApplyTexture");
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
    if (portable::IsTextureHandle(static_cast<std::uint64_t>(handle))) {
        RequireInitialised();
        RequireOk(BB_ReleaseTexture(static_cast<BB_Handle>(handle)), "BB_ReleaseTexture");
        return;
    }
    if (textures_.erase(handle) == 0) {
        throw Error(BB_E_TEXTURE_NOT_FOUND, kTextureNotFoundMessage);
    }
}

void Engine::ClearTextures() noexcept {
    textures_.clear();
    BB_ClearTextures();
}

long Engine::TextureCount() const {
    return static_cast<long>(textures_.size()) + static_cast<long>(BB_GetTextureCount());
}
} // namespace bb
