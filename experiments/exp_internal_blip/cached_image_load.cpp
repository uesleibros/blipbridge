/**
 * @file cached_image_load.cpp
 * Creates an Office cached image from bytes already in memory, then releases it.
 *
 * This is the smallest possible native step toward a reusable texture: it proves
 * that byte input can reach Office's decoder through a standard `IStream` with
 * no file on disk, and it touches no document state whatsoever. No Shape, no
 * slide, no presentation and no property record is involved, so the experiment
 * cannot corrupt a document; the worst case is a failed call.
 *
 * ## What is called, and why it is not a private offset
 *
 * `GFX.DLL` exports the creator by name at ordinal 236:
 *
 * ```text
 * ?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@
 *     AEAV?$TCntPtr@UIImage@GEL@@@4@ PEAUIStream@@
 *     W4IStreamCopyInstruction@12@ PEBVMD4UID@4@ _N @Z
 * ```
 *
 * so it is resolved with `GetProcAddress`, not with a hard-coded RVA.
 *
 * ## Why the calling convention below is believed correct
 *
 * The name mangling gives the parameter list. `GFX +0x7680` begins
 * `mov %rcx,0x8(%rsp)` and ends `mov %rbx,%rax`, which is the MSVC x64
 * convention for returning a non-trivial class: RCX is a hidden pointer to the
 * caller's return storage, and it is also the return value. The remaining
 * arguments therefore shift one register right, which is exactly what Office's
 * own call site at `OART +0x8F54E` does:
 *
 * ```text
 * rcx        = &TCntPtr<ICachedImage>   (return storage)
 * rdx        = &TCntPtr<IImage>         (in/out, empty on entry)
 * r8         = IStream*
 * r9d        = 0                        (IStreamCopyInstruction)
 * [rsp+0x20] = const MD4UID*            (a 16-byte local)
 * [rsp+0x28] = false                    (bool)
 * ```
 *
 * The wrapper also settles ownership: it reads the worker's result and zeroes
 * the source rather than AddRef'ing, so both returned pointers carry one
 * reference each that this function must release. Release is vtable slot +8;
 * `GFX +0xA1A0` decrements a 32-bit count at object+8 and destroys at 1.
 *
 * See docs/record_construction.md for the derivation and docs/resource_lifetime.md
 * for the intrusive counting rules.
 */

#include "../experiment_api.hpp"

#include <blipbridge/dispatch.hpp>

#include <cstdint>
#include <cstring>
#include <sstream>
#include <vector>

namespace {

/// Only this exact Office build has been observed and validated.
constexpr DWORD kSupportedVersionHigh = 0x00100000;   // 16.0
constexpr DWORD kSupportedVersionLow = (14334u << 16) | 20848u;

/// Exported symbol; resolved by name, never by address.
constexpr char kCreateFromStreamSymbol[] =
    "?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@"
    "AEAV?$TCntPtr@UIImage@GEL@@@4@PEAUIStream@@W4IStreamCopyInstruction@12@"
    "PEBVMD4UID@4@_N@Z";

/// Vtables of the two objects the creator produces, used only to verify.
constexpr uintptr_t kCachedImageVtableRva = 0x409DC0;   // gfx.dll
constexpr uintptr_t kImageVtableRva = 0x4055C8;         // gfx.dll

/// Intrusive layout shared by both: 32-bit count at +8, Release at vtable +8.
constexpr size_t kCountOffset = 8;
constexpr size_t kReleaseSlotOffset = 8;

/// Values Office passes at OART +0x8F54E.
constexpr int kStreamCopyInstruction = 0;
constexpr bool kCreateFlag = false;
constexpr size_t kUidSize = 16;

/**
 * `Ofc::TCntPtr<T>` as this code needs it: one pointer. It is only ever used as
 * storage the creator writes into, so no copy semantics are modelled.
 */
struct CountedPointer {
    void* value = nullptr;
};

/// MSVC x64 signature with the hidden return pointer written out explicitly.
using CreateCachedImageFromStream = CountedPointer*(__stdcall*)(
    CountedPointer* returnStorage,
    CountedPointer* imageInOut,
    IStream* stream,
    int copyInstruction,
    const void* uid,
    bool flag);

/**
 * Releases one intrusive reference through vtable slot +8.
 * Never read the object after this returns; the count may have reached zero.
 */
void ReleaseIntrusive(void* object) noexcept {
    if (!object) {
        return;
    }
    auto vtable = *reinterpret_cast<void* const* const*>(object);
    using Release = void(__stdcall*)(void*);
    auto release = reinterpret_cast<Release>(
        *reinterpret_cast<void* const*>(
            reinterpret_cast<const std::uint8_t*>(vtable) + kReleaseSlotOffset));
    release(object);
}

std::uint32_t IntrusiveCount(const void* object) {
    std::uint32_t count = 0;
    std::memcpy(&count, reinterpret_cast<const std::uint8_t*>(object) + kCountOffset,
                sizeof(count));
    return count;
}

/// Owns one intrusive reference for the duration of the experiment.
class CountedReference {
public:
    CountedReference() = default;
    CountedReference(const CountedReference&) = delete;
    CountedReference& operator=(const CountedReference&) = delete;
    ~CountedReference() { ReleaseIntrusive(storage.value); }

    CountedPointer storage;
};

/// Throws unless @p moduleName is loaded and is the build these offsets came from.
HMODULE RequireSupportedModule(const wchar_t* moduleName, const char* description) {
    HMODULE module = GetModuleHandleW(moduleName);
    if (!module) {
        throw bb::Error(E_NOTIMPL,
                        std::string(description) +
                            " is not loaded; warm it with an ordinary picture fill first");
    }
    wchar_t modulePath[MAX_PATH * 4]{};
    if (!GetModuleFileNameW(module, modulePath, static_cast<DWORD>(std::size(modulePath)))) {
        throw bb::Error(E_NOTIMPL, std::string("Cannot resolve the ") + description + " path");
    }
    DWORD ignored = 0;
    const DWORD versionSize = GetFileVersionInfoSizeW(modulePath, &ignored);
    std::vector<BYTE> version(versionSize);
    if (!versionSize || !GetFileVersionInfoW(modulePath, 0, versionSize, version.data())) {
        throw bb::Error(E_NOTIMPL,
                        std::string("Cannot read the ") + description + " version resource");
    }
    VS_FIXEDFILEINFO* information = nullptr;
    UINT length = 0;
    if (!VerQueryValueW(version.data(), L"\\", reinterpret_cast<void**>(&information), &length) ||
        length < sizeof(*information)) {
        throw bb::Error(E_NOTIMPL,
                        std::string("Cannot read the ") + description + " version information");
    }
    if (information->dwFileVersionMS != kSupportedVersionHigh ||
        information->dwFileVersionLS != kSupportedVersionLow) {
        throw bb::Error(E_NOTIMPL,
                        "Cached image loading supports only Office 16.0.14334.20848");
    }
    return module;
}

/// Releases the HGLOBAL-backed stream and its memory together.
struct OwnedStream {
    IStream* value = nullptr;
    ~OwnedStream() {
        if (value) {
            value->Release();
        }
    }
};

OwnedStream CreateStreamOverBytes(SAFEARRAY* bytes) {
    if (!bytes || SafeArrayGetDim(bytes) != 1) {
        throw bb::Error(E_INVALIDARG, "Expected a one-dimensional Byte array");
    }
    LONG lower = 0;
    LONG upper = 0;
    bb::check(SafeArrayGetLBound(bytes, 1, &lower), "SafeArrayGetLBound");
    bb::check(SafeArrayGetUBound(bytes, 1, &upper), "SafeArrayGetUBound");
    if (upper < lower) {
        throw bb::Error(E_INVALIDARG, "Empty image bytes");
    }
    const size_t size = static_cast<size_t>(upper - lower) + 1;

    // GMEM_MOVEABLE is required by CreateStreamOnHGlobal; passing TRUE makes the
    // stream free it, so there is a single owner from here on.
    HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, size);
    if (!memory) {
        throw std::bad_alloc();
    }
    void* destination = GlobalLock(memory);
    if (!destination) {
        GlobalFree(memory);
        throw bb::Error(E_OUTOFMEMORY, "Cannot lock image storage");
    }
    void* source = nullptr;
    const HRESULT accessed = SafeArrayAccessData(bytes, &source);
    if (SUCCEEDED(accessed)) {
        std::memcpy(destination, source, size);
        SafeArrayUnaccessData(bytes);
    }
    GlobalUnlock(memory);
    if (FAILED(accessed)) {
        GlobalFree(memory);
        bb::check(accessed, "SafeArrayAccessData");
    }

    OwnedStream stream;
    const HRESULT created = CreateStreamOnHGlobal(memory, TRUE, &stream.value);
    if (FAILED(created)) {
        GlobalFree(memory);
        bb::check(created, "CreateStreamOnHGlobal");
    }
    return stream;
}

} // namespace

/**
 * Decodes @p bytes into an Office cached image and releases it again, reporting
 * what was observed.
 *
 * Ownership: the creator hands back one reference on each of the cached image
 * and the image; both are released here before returning, in the reverse of the
 * order Office's own caller uses. Nothing is retained, and no document object is
 * read or written.
 */
std::wstring loadCachedImageExperiment(SAFEARRAY* bytes) {
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "Cached image loading requires the PowerPoint host");
    }
    const HMODULE gfx = RequireSupportedModule(L"gfx.dll", "GFX");
    const auto gfxBase = reinterpret_cast<std::uintptr_t>(gfx);
    auto create = reinterpret_cast<CreateCachedImageFromStream>(
        reinterpret_cast<void*>(GetProcAddress(gfx, kCreateFromStreamSymbol)));
    if (!create) {
        throw bb::Error(E_NOTIMPL, "GFX does not export the stream cached-image creator");
    }

    OwnedStream stream = CreateStreamOverBytes(bytes);

    // Office passes a 16-byte identifier local; its contents were not recovered
    // statically, so pass zeroes, which is what the record slot also carries.
    const std::uint8_t uid[kUidSize]{};

    CountedReference cached;
    CountedReference image;
    create(&cached.storage, &image.storage, stream.value, kStreamCopyInstruction, uid,
           kCreateFlag);

    if (!cached.storage.value) {
        throw bb::Error(E_FAIL, "Creator returned no cached image for these bytes");
    }
    const auto cachedVtable = *reinterpret_cast<std::uintptr_t*>(cached.storage.value);
    if (cachedVtable != gfxBase + kCachedImageVtableRva) {
        // Do not release something whose layout was not confirmed.
        cached.storage.value = nullptr;
        image.storage.value = nullptr;
        throw bb::Error(E_NOTIMPL, "Cached image vtable does not match the validated layout");
    }

    std::wostringstream out;
    out << L"cached=0x" << std::hex << reinterpret_cast<std::uintptr_t>(cached.storage.value)
        << std::dec << L";cachedCount=" << IntrusiveCount(cached.storage.value) << L';';
    if (image.storage.value) {
        const auto imageVtable = *reinterpret_cast<std::uintptr_t*>(image.storage.value);
        out << L"image=0x" << std::hex << reinterpret_cast<std::uintptr_t>(image.storage.value)
            << std::dec << L";imageCount=" << IntrusiveCount(image.storage.value)
            << L";imageVtableMatches=" << (imageVtable == gfxBase + kImageVtableRva ? 1 : 0)
            << L';';
        if (imageVtable != gfxBase + kImageVtableRva) {
            image.storage.value = nullptr;   // unknown layout: leak rather than corrupt
        }
    } else {
        out << L"image=none;";
    }
    out << L"fileFree=1;";
    return out.str();
}
