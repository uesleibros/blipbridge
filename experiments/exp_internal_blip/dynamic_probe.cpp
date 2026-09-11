/**
 * @file dynamic_probe.cpp
 * Research: is there any way to change the pixels behind an image a Shape is
 * already showing, without applying a new fill?
 *
 * The whole dynamic-texture idea rests on one question that can be answered
 * cheaply before any reverse engineering is attempted: when
 * `GEL::ICachedImage::Create` is handed a raw pixel buffer, does it **copy**
 * those bytes or **reference** them?
 *
 * If it references the caller's memory, a dynamic texture is nearly free - write
 * to the buffer, ask Office to repaint, done, with the same image identity and
 * no fill transaction. If it copies, then the pixels a Shape displays live
 * somewhere inside GFX, and reaching them means finding and writing that storage,
 * which is a different and far more dangerous project.
 *
 * So this holds the source buffer alive, applies it once, and then mutates it.
 * The harness exports the Shape before and after and compares the renders. No
 * assumption, no disassembly: the answer is whatever PowerPoint draws.
 *
 * Research only, and deliberately so - nothing here is a candidate for the
 * public API until the question above has an answer worth building on.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/native_texture.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <sstream>
#include <vector>

namespace {

/**
 * The buffer handed to GFX, kept alive between calls.
 *
 * Deliberately process-lifetime: the point of the probe is to still own the
 * memory after Office has taken the image, so that writing to it can be observed
 * - or not - in what gets drawn. Freeing it would answer a different question.
 */
struct DynamicProbe {
    std::vector<std::uint8_t> pixels;
    bb::office::TextureRef texture;
    std::uint32_t width = 0;
    std::uint32_t height = 0;

    static DynamicProbe& Instance() {
        static DynamicProbe probe;
        return probe;
    }
};

/// Fills the buffer with one colour, so a render is unambiguous.
void Paint(std::vector<std::uint8_t>& pixels, std::uint8_t b, std::uint8_t g, std::uint8_t r) {
    for (std::size_t index = 0; index + 3 < pixels.size(); index += 4) {
        pixels[index + 0] = b;
        pixels[index + 1] = g;
        pixels[index + 2] = r;
        pixels[index + 3] = 255;
    }
}

} // namespace

std::wstring probeDynamicTexture(IDispatch* shape, long step) {
    if (!shape) {
        throw bb::Error(E_POINTER, "Missing Shape");
    }
    DynamicProbe& probe = DynamicProbe::Instance();
    std::wostringstream out;

    switch (step) {
    case 0: {
        // Create a solid red image from a buffer we keep, and apply it once.
        probe.width = 64;
        probe.height = 64;
        probe.pixels.assign(static_cast<std::size_t>(probe.width) * probe.height * 4, 0);
        Paint(probe.pixels, 0, 0, 255);

        probe.texture = bb::office::CreateTextureFromPixels(
            probe.pixels.data(), probe.width, probe.height,
            static_cast<std::int32_t>(probe.width) * 4);
        bb::Value fill = bb::get(shape, L"Fill");
        bb::office::ApplyTextureRef(fill.obj(), probe.texture);

        out << L"step=0;applied=1;imageId=" << bb::office::TextureIdOf(probe.texture)
            << L";cached=0x" << std::hex
            << reinterpret_cast<std::uintptr_t>(bb::office::CachedImageOf(probe.texture))
            << std::dec << L';';
        break;
    }
    case 1: {
        // Overwrite the same buffer with blue. No new image, no apply, no fill
        // transaction. If GFX referenced this memory, the Shape is now blue.
        if (probe.pixels.empty()) {
            throw bb::Error(E_FAIL, "Run step 0 first");
        }
        Paint(probe.pixels, 255, 0, 0);
        out << L"step=1;mutated=1;imageId=" << bb::office::TextureIdOf(probe.texture)
            << L";cached=0x" << std::hex
            << reinterpret_cast<std::uintptr_t>(bb::office::CachedImageOf(probe.texture))
            << std::dec << L';';
        break;
    }
    case 2: {
        /*
         * Re-apply the *same* cached image, to separate two possibilities that
         * step 1 cannot tell apart on its own: GFX copied the pixels at creation
         * time, or GFX referenced them but Office is drawing from a cache of its
         * own that nothing has invalidated.
         *
         * If the buffer is referenced and only the repaint was missing, this
         * makes the Shape blue - and the image identity is unchanged, which
         * would be most of a dynamic texture. If it stays red, the bytes were
         * copied when the image was created and no amount of invalidation will
         * reach them.
         */
        if (!probe.texture) {
            throw bb::Error(E_FAIL, "Run step 0 first");
        }
        bb::Value fill = bb::get(shape, L"Fill");
        bb::office::ApplyTextureRef(fill.obj(), probe.texture);
        out << L"step=2;reapplied=1;imageId=" << bb::office::TextureIdOf(probe.texture) << L';';
        break;
    }
    case 3: {
        // Release, so a run leaves nothing behind.
        probe.texture.reset();
        probe.pixels.clear();
        probe.pixels.shrink_to_fit();
        out << L"step=3;released=1;";
        break;
    }
    default:
        throw bb::Error(E_INVALIDARG, "Steps are 0 create, 1 mutate, 2 re-apply, 3 release");
    }
    return out.str();
}
