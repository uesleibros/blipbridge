/**
 * @file pixel_texture.cpp
 * Research: creating a cached image straight from raw pixels.
 *
 * GFX exports a second cached-image creator that takes a pixel buffer instead of
 * a stream:
 *
 * ```text
 * ?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@
 *     AEAV?$TCntPtr@UIImage@GEL@@@4@ PEBX I I H
 *     W4SurfaceFormat@ARC@@
 *     AEBU?$TVector2@V?$TUnits@MU?$TUnitsRatioTag@UDevicePixels@Math@@UInches@2@@Math@@@Math@@@Math@@
 * @Z
 * ```
 *
 * which is
 *
 * ```c
 * TCntPtr<ICachedImage> Create(TCntPtr<IImage>& image, const void* pixels,
 *                              unsigned width, unsigned height, int stride,
 *                              ARC::SurfaceFormat format,
 *                              const Math::TVector2<float>& dpi);
 * ```
 *
 * If this works it removes PNG/JPEG encoding and decoding entirely for callers
 * that already hold pixels.
 *
 * ## What is known and what is not
 *
 * The calling convention is settled by the prologue at GFX +0x193C00, which is
 * the same shape as the stream creator: RCX is the hidden sret pointer, the
 * result is *moved* into it rather than AddRef'd - so the returned pointer
 * carries exactly one reference - and a local temporary is released through
 * GFX +0xA1A0 on the way out. The two `unsigned` parameters are stored into
 * adjacent stack slots and passed on by address, which is consistent with a
 * width/height pair.
 *
 * `ARC::SurfaceFormat` is **not** known. It is an enum with no symbols, so this
 * experiment exists to find which value means which channel order by creating a
 * texture of a known colour and looking at what PowerPoint renders. Nothing here
 * assumes a value; the caller supplies one and the harness reports what came out.
 */

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../experiment_api.hpp"

#include <blipbridge/dispatch.hpp>
#include <cstring>
#include <sstream>

namespace {

constexpr char kCreateFromPixelsSymbol[] =
    "?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@"
    "AEAV?$TCntPtr@UIImage@GEL@@@4@PEBXIIHW4SurfaceFormat@ARC@@"
    "AEBU?$TVector2@V?$TUnits@MU?$TUnitsRatioTag@UDevicePixels@Math@@UInches@2@"
    "@Math@@@Math@@@Math@@@Z";

constexpr std::uintptr_t kCachedImageVtableRva = 0x409DC0; // gfx.dll
constexpr std::uintptr_t kImageVtableRva = 0x4055C8;       // gfx.dll

/// `Ofc::TCntPtr<T>` as this code needs it: one pointer of storage.
struct CountedPointer {
    void* value = nullptr;
};

/// `Math::TVector2<float>` for the DPI argument: two floats, x then y.
struct Vector2 {
    float x = 96.0f;
    float y = 96.0f;
};

using CreateCachedImageFromPixels = CountedPointer*(__stdcall*)(CountedPointer * returnStorage,
                                                                CountedPointer* imageInOut,
                                                                const void* pixels,
                                                                unsigned int width,
                                                                unsigned int height,
                                                                int stride,
                                                                int surfaceFormat,
                                                                const Vector2* dpi);

/// Releases one intrusive GFX reference when it goes out of scope.
class CountedReference {
  public:
    CountedReference() = default;
    CountedReference(const CountedReference&) = delete;
    CountedReference& operator=(const CountedReference&) = delete;

    ~CountedReference() {
        bb::oart::ReleaseIntrusive(storage.value);
    }

    CountedPointer storage;
};

} // namespace

/**
 * Creates a cached image from raw pixels and applies it to @p fill.
 *
 * Everything the format probe needs is a parameter, because none of it is
 * known: the caller chooses the surface-format value and the harness reports
 * what PowerPoint made of it. The bytes are borrowed for the duration of the
 * call - the creator either copies them or takes its own reference before
 * returning, and this function releases everything it created before it exits.
 *
 * Returns a field list; throws bb::Error if creation or the apply fails, so a
 * rejected format value is distinguishable from an accepted one.
 */
std::wstring pixelTextureExperiment(
    IDispatch* fill, SAFEARRAY* pixels, long width, long height, long stride, long surfaceFormat) {
    if (!pixels || SafeArrayGetDim(pixels) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional Byte array of pixels");
    }
    if (width <= 0 || height <= 0) {
        throw bb::Error(E_INVALIDARG, "Width and height must be positive");
    }
    LONG lower = 0;
    LONG upper = 0;
    bb::check(SafeArrayGetLBound(pixels, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(pixels, 1, &upper), "SafeArrayGetUBound");
    const long long supplied = static_cast<long long>(upper) - lower + 1;
    const long long needed = static_cast<long long>(stride) * height;
    if (supplied < needed) {
        std::ostringstream out;
        out << "Pixel buffer is " << supplied << " bytes but stride*height is " << needed;
        throw bb::Error(E_INVALIDARG, out.str());
    }

    const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fill);
    const bb::oart::ApplyFunctions functions = bb::oart::ResolveApplyFunctions(target.oartBase);

    const HMODULE gfx = bb::oart::RequireSupportedModule(L"gfx.dll", "GFX");
    const auto gfxBase = reinterpret_cast<std::uintptr_t>(gfx);
    auto create = reinterpret_cast<CreateCachedImageFromPixels>(
        reinterpret_cast<void*>(GetProcAddress(gfx, kCreateFromPixelsSymbol)));
    if (!create) {
        throw bb::Error(E_NOTIMPL, "GFX does not export the raw-pixel cached-image creator");
    }

    void* raw = nullptr;
    bb::check(SafeArrayAccessData(pixels, &raw), "SafeArrayAccessData");

    struct Unaccess {
        SAFEARRAY* value;

        ~Unaccess() {
            SafeArrayUnaccessData(value);
        }
    } unaccess{pixels};

    const Vector2 dpi; // 96 dpi in both axes, the ordinary screen default
    CountedReference cached;
    CountedReference image;
    create(&cached.storage,
           &image.storage,
           raw,
           static_cast<unsigned int>(width),
           static_cast<unsigned int>(height),
           static_cast<int>(stride),
           static_cast<int>(surfaceFormat),
           &dpi);

    if (!cached.storage.value) {
        std::ostringstream out;
        out << "Surface format " << surfaceFormat << " produced no cached image";
        throw bb::Error(E_FAIL, out.str());
    }
    const std::uintptr_t cachedVtable = bb::oart::LoadPointer(cached.storage.value, 0);
    if (cachedVtable != gfxBase + kCachedImageVtableRva) {
        // Unknown layout: drop rather than release, so a wrong guess leaks a
        // little instead of corrupting the allocator.
        cached.storage.value = nullptr;
        image.storage.value = nullptr;
        throw bb::Error(E_NOTIMPL, "Cached image vtable does not match the validated layout");
    }

    std::wostringstream out;
    out << L"format=" << surfaceFormat << L";width=" << width << L";height=" << height
        << L";stride=" << stride << L";cached=0x" << std::hex
        << reinterpret_cast<std::uintptr_t>(cached.storage.value) << std::dec << L";cachedCount="
        << bb::oart::IntrusiveCount(cached.storage.value) << L';';
    if (image.storage.value) {
        const std::uintptr_t imageVtable = bb::oart::LoadPointer(image.storage.value, 0);
        out << L"imageVtableMatches=" << (imageVtable == gfxBase + kImageVtableRva ? 1 : 0)
            << L";imageCount=" << bb::oart::IntrusiveCount(image.storage.value) << L';';
        if (imageVtable != gfxBase + kImageVtableRva) {
            image.storage.value = nullptr;
        }
    } else {
        out << L"image=none;";
    }

    // The apply path is entirely unchanged: a cached image is a cached image,
    // wherever it came from.
    bb::oart::ApplyCachedImage(functions, target, cached.storage.value);
    out << L"applied=1;";
    return out.str();
}
