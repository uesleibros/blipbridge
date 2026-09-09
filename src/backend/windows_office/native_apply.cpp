/**
 * @file native_apply.cpp
 * Applies a picture fill to an existing normal Shape without `Fill.UserPicture`.
 *
 * The pipeline, all of it in this one call:
 *
 * ```text
 * Byte[] -> HGLOBAL IStream -> GEL::ICachedImage::Create (a GFX export)
 *        -> OART image sub-record -> OART property record
 *        -> OART transaction -> receiver vtable +0x78 -> fill applied
 *        -> destroy transaction, sub-record, record -> release our references
 * ```
 *
 * No donor Shape, no `PickUp`/`Apply`, no `UserPicture`, and no file: the bytes
 * reach Office's decoder through a memory stream this function creates.
 *
 * ## Safety rules this file follows
 *
 *  - Every private OART function is byte-verified at its RVA before it is
 *    called; a single mismatched byte aborts before anything is constructed.
 *  - The receiver is resolved immediately before use and never cached, because
 *    a deleted Shape still passes every pointer and vtable check.
 *  - Each constructed Office object is owned by a scope guard, so the three
 *    destructors run in the handler's own reverse order even if a later step
 *    throws.
 *  - Reference counts are sampled at every stage. An unexplained delta is
 *    reported rather than hidden.
 *
 * ## Why the record is built this way
 *
 * `OART +0x89C860`, the real `UserPicture` handler, builds the whole ~0x4E8-byte
 * record with Office's own constructors and then sets only a few fields. This
 * reproduces that sequence, with two deliberate differences:
 *
 *  - The file loader `OART +0x950070` is replaced by the exported cached-image
 *    creator plus `OART +0x8F94C`, which is exactly what the loader does
 *    internally once it has an `IStream`.
 *  - The counted 16-byte slot at record+0x2A0 is set exactly as the handler sets
 *    it. That slot is what separates a picture fill from a texture fill: the
 *    FillFormat vtable's next entry, `OART +0x8A1530` (UserTextured), calls the
 *    *same* handler `+0x89C860` and differs only in building this argument with
 *    `+0x239950` instead of `+0x158C40`. Leaving the slot unset was measured to
 *    produce `Fill.Type = 4` (textured) instead of `6` (picture).
 *
 * See docs/oart_abi.md for the per-function ABI table and its evidence.
 */

#include "native_apply.hpp"
#include "oart_layout.hpp"

#include <blipbridge/dispatch.hpp>

#include <cstring>
#include <sstream>
#include <string>
#include <vector>

namespace {

using bb::oart::FillTarget;
using bb::oart::GuardedFunction;
using bb::oart::MakeGuarded;

// -- verified private entry points -----------------------------------------
// Signature bytes were recorded on OART 16.0.14334.20848 x64 at the same time
// each function's ABI was derived. See docs/oart_abi.md.

constexpr std::uint8_t kRecordConstructorBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x08, 0x48, 0x89, 0x74, 0x24, 0x10,
    0x57, 0x48, 0x83, 0xec, 0x30, 0x48};
constexpr std::uint8_t kClearSlotsBytes[] = {
    0xba, 0xfa, 0xff, 0xff, 0xff, 0x21, 0x11, 0x21, 0x51, 0x08,
    0x21, 0x51, 0x28, 0x8d, 0x42, 0x08};
constexpr std::uint8_t kImageRecordConstructorBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x08, 0x48, 0x89, 0x6c, 0x24, 0x10,
    0x48, 0x89, 0x74, 0x24, 0x18, 0x57};
constexpr std::uint8_t kInstallCachedImageBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x20, 0x56, 0x57, 0x41, 0x56, 0x48,
    0x83, 0xec, 0x30, 0x83, 0x64, 0x24};
constexpr std::uint8_t kTransferImageSlotBytes[] = {
    0x40, 0x53, 0x48, 0x83, 0xec, 0x20, 0x48, 0x8b, 0xd9, 0x48,
    0x83, 0xc1, 0x08, 0xe8, 0x6a, 0x4b};
constexpr std::uint8_t kTransactionConstructorBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x10, 0x48, 0x89, 0x6c, 0x24, 0x18,
    0x48, 0x89, 0x4c, 0x24, 0x08, 0x56};
constexpr std::uint8_t kRecordDestructorBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x08, 0x48, 0x89, 0x74, 0x24, 0x10,
    0x57, 0x48, 0x83, 0xec, 0x20, 0x48};
constexpr std::uint8_t kImageRecordDestructorBytes[] = {
    0x48, 0x89, 0x4c, 0x24, 0x08, 0x53, 0x56, 0x57, 0x41, 0x54,
    0x41, 0x55, 0x41, 0x56, 0x41, 0x57};
constexpr std::uint8_t kTransactionDestructorBytes[] = {
    0x40, 0x53, 0x48, 0x83, 0xec, 0x20, 0x48, 0x8b, 0xd9, 0x48,
    0x83, 0xc1, 0x18, 0xe8, 0xe6, 0xeb};
constexpr std::uint8_t kStretchHolderBuilderBytes[] = {
    0x48, 0x89, 0x5c, 0x24, 0x08, 0x57, 0x48, 0x83, 0xec, 0x30,
    0x48, 0x8b, 0xf9, 0x48, 0x8d, 0x05};
constexpr std::uint8_t kSetCountedSlotBytes[] = {
    0x40, 0x53, 0x48, 0x83, 0xec, 0x20, 0x48, 0x8b, 0xd9, 0x48,
    0x83, 0xc1, 0x08, 0xe8, 0x7e, 0xec};
constexpr std::uint8_t kHolderDestructorBytes[] = {
    0x48, 0x83, 0xec, 0x28, 0x48, 0x8b, 0x41, 0x08, 0x48, 0x83,
    0xf8, 0x01, 0x77, 0x05, 0x48, 0x83};

constexpr GuardedFunction kRecordConstructor =
    MakeGuarded(0x14110, "property record constructor", kRecordConstructorBytes);
constexpr GuardedFunction kClearSlots =
    MakeGuarded(0x14F7B0, "property slot clear", kClearSlotsBytes);
constexpr GuardedFunction kImageRecordConstructor =
    MakeGuarded(0x14580, "image sub-record constructor", kImageRecordConstructorBytes);
constexpr GuardedFunction kInstallCachedImage =
    MakeGuarded(0x8F94C, "cached image install", kInstallCachedImageBytes);
constexpr GuardedFunction kTransferImageSlot =
    MakeGuarded(0x22BCF4, "image slot transfer", kTransferImageSlotBytes);
constexpr GuardedFunction kTransactionConstructor =
    MakeGuarded(0x48870, "transaction constructor", kTransactionConstructorBytes);
constexpr GuardedFunction kRecordDestructor =
    MakeGuarded(0xAF80, "property record destructor", kRecordDestructorBytes);
constexpr GuardedFunction kImageRecordDestructor =
    MakeGuarded(0x3D870, "image sub-record destructor", kImageRecordDestructorBytes);
constexpr GuardedFunction kTransactionDestructor =
    MakeGuarded(0x8C388, "transaction destructor", kTransactionDestructorBytes);
constexpr GuardedFunction kStretchHolderBuilder =
    MakeGuarded(0x158C40, "stretch holder builder", kStretchHolderBuilderBytes);
constexpr GuardedFunction kSetCountedSlot =
    MakeGuarded(0x15AC70, "counted slot setter", kSetCountedSlotBytes);
constexpr GuardedFunction kHolderDestructor =
    MakeGuarded(0xB210, "counted holder destructor", kHolderDestructorBytes);

// -- object sizes ----------------------------------------------------------
// Constructor write ranges and one heap allocation site give these; the buffers
// below add margin because over-allocating a local costs nothing and a short
// buffer would be memory corruption.
constexpr std::size_t kRecordSize = 0x4E8;          // +0x14110 writes through +0x4E4
constexpr std::size_t kRecordBufferSize = 0x520;
constexpr std::size_t kImageRecordSize = 0x200;     // heap allocation at OART +0xD9A62
constexpr std::size_t kImageRecordBufferSize = 0x220;
constexpr std::size_t kTransactionSize = 0x510;     // +0x48870 writes through +0x508
constexpr std::size_t kTransactionBufferSize = 0x600;

// -- record field layout ---------------------------------------------------
constexpr std::size_t kFillKindSlotOffset = 0x00;
constexpr std::size_t kFillKindPayloadOffset = 0x04;
constexpr std::uint32_t kPictureFillKind = 3;
constexpr std::size_t kImageSlotOffset = 0x88;
constexpr std::size_t kImageSubRecordOffset = 0x90;
constexpr std::size_t kCachedImageInSubRecordOffset = 0xF0;
constexpr std::size_t kTrailingFlagSlotOffset = 0x4D8;
constexpr std::size_t kTrailingFlagPayloadOffset = 0x4DC;
/// Stretch-versus-tile slot; see the file comment for how that was established.
constexpr std::size_t kStretchSlotOffset = 0x2A0;
constexpr std::size_t kStretchValueSize = 16;

/// Slot discriminator idioms, both taken verbatim from the handler.
constexpr std::uint32_t MarkSlotSet(std::uint32_t value) {
    return (value & ~6u) | 1u;
}

// -- transaction arguments -------------------------------------------------
// The handler passes these literally at OART +0x89CA2E..+0x89CA48.
constexpr std::uint32_t kTransactionFlags = 0;
constexpr std::uint32_t kPictureFillIdentifier = 0xA042008E;

/// Receiver vtable slot the handler calls with the finished transaction.
constexpr std::size_t kApplyTransactionSlot = 0x78;

// The shared function-pointer types live in native_apply.hpp so the texture
// store uses exactly the same declarations.
using bb::oart::ApplyFunctions;
using bb::oart::BuildStretchHolder;
using bb::oart::ClearSlots;
using bb::oart::Destructor;
using bb::oart::ImageRecordConstructor;
using bb::oart::InstallCachedImage;
using bb::oart::RecordConstructor;
using bb::oart::SetCountedSlot;
using bb::oart::TransactionConstructor;
using bb::oart::TransferImageSlot;

/// The receiver method the handler calls with the finished transaction.
using ApplyTransaction = void*(__stdcall*)(void* receiver, void* transaction);

/**
 * The `{payload, descriptor}` pair `OART +0x158C40` fills in and `OART +0xB210`
 * destroys. Both tolerate a zeroed pair, which is why this starts zeroed: the
 * builder destroys whatever was there before overwriting it.
 */
struct CountedHolder {
    void* payload = nullptr;
    const void* descriptor = nullptr;
};

/// Destroys a CountedHolder through Office's own destructor on scope exit.
class HolderGuard {
public:
    HolderGuard(CountedHolder* holder, Destructor destructor)
        : holder_(holder), destructor_(destructor) {}
    HolderGuard(const HolderGuard&) = delete;
    HolderGuard& operator=(const HolderGuard&) = delete;
    ~HolderGuard() { Destroy(); }

    void Destroy() {
        if (holder_ && destructor_) {
            Destructor destructor = destructor_;
            void* holder = holder_;
            holder_ = nullptr;
            destructor_ = nullptr;
            destructor(holder);
        }
    }

private:
    CountedHolder* holder_ = nullptr;
    Destructor destructor_ = nullptr;
};

/// `Ofc::TCntPtr<T>` as this code needs it: one pointer of storage.
struct CountedPointer {
    void* value = nullptr;
};

using CreateCachedImageFromStream = CountedPointer*(__stdcall*)(
    CountedPointer* returnStorage, CountedPointer* imageInOut, IStream* stream,
    int copyInstruction, const void* uid, bool flag);

constexpr char kCreateFromPixelsSymbol[] =
    "?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@"
    "AEAV?$TCntPtr@UIImage@GEL@@@4@PEBXIIHW4SurfaceFormat@ARC@@"
    "AEBU?$TVector2@V?$TUnits@MU?$TUnitsRatioTag@UDevicePixels@Math@@UInches@2@"
    "@Math@@@Math@@@Math@@@Z";

/// ARC::SurfaceFormat value passed to the raw-pixel creator. Probing 0..24 gave
/// an identical correct BGRA render for every value, so this is a fixed choice
/// rather than a meaningful parameter; see docs/pixel_textures.md.
constexpr int kSurfaceFormatBgra32 = 0;

/// Math::TVector2<float> for the DPI argument: 96 dpi in both axes.
struct Vector2 {
    float x = 96.0f;
    float y = 96.0f;
};

using CreateCachedImageFromPixelBuffer = CountedPointer*(__stdcall*)(
    CountedPointer* returnStorage, CountedPointer* imageInOut, const void* pixels,
    unsigned int width, unsigned int height, int stride, int surfaceFormat,
    const Vector2* dpi);

constexpr char kCreateFromStreamSymbol[] =
    "?Create@ICachedImage@GEL@@SA?AV?$TCntPtr@UICachedImage@GEL@@@Ofc@@"
    "AEAV?$TCntPtr@UIImage@GEL@@@4@PEAUIStream@@W4IStreamCopyInstruction@12@"
    "PEBVMD4UID@4@_N@Z";

constexpr std::uintptr_t kCachedImageVtableRva = 0x409DC0;   // gfx.dll
constexpr int kStreamCopyInstruction = 0;
constexpr bool kCreateFlag = false;
constexpr std::size_t kUidSize = 16;

/// Releases one intrusive GFX reference when it goes out of scope.
class CountedReference {
public:
    CountedReference() = default;
    CountedReference(const CountedReference&) = delete;
    CountedReference& operator=(const CountedReference&) = delete;
    ~CountedReference() { bb::oart::ReleaseIntrusive(storage.value); }

    CountedPointer storage;
};

/**
 * Runs an Office destructor on a constructed local when the scope ends.
 * The three of these in the apply below unwind in the same order the real
 * handler uses: transaction, image sub-record, property record.
 */
class ConstructedObject {
public:
    ConstructedObject(void* object, Destructor destructor)
        : object_(object), destructor_(destructor) {}
    ConstructedObject(const ConstructedObject&) = delete;
    ConstructedObject& operator=(const ConstructedObject&) = delete;
    ~ConstructedObject() { Destroy(); }

    /// Runs the destructor once. Safe to call explicitly and then again on scope
    /// exit, which is how the success path destroys in a chosen order while the
    /// failure path still unwinds correctly.
    void Destroy() {
        if (object_ && destructor_) {
            Destructor destructor = destructor_;
            void* object = object_;
            object_ = nullptr;
            destructor_ = nullptr;
            destructor(object);
        }
    }

private:
    void* object_ = nullptr;
    Destructor destructor_ = nullptr;
};

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

namespace bb::oart {

ApplyFunctions ResolveApplyFunctions(std::uintptr_t oart) {
    // Byte-verifying twelve entry points costs a VirtualQuery and a memcmp each.
    // The bytes at a given RVA cannot change while the image stays loaded, so
    // verify once per base address; a reloaded OART lands on a different base and
    // is verified again. Nothing document-derived is cached here.
    static std::uintptr_t verifiedBase = 0;
    static ApplyFunctions verified;
    if (verifiedBase == oart) {
        return verified;
    }

    ApplyFunctions functions;
    functions.constructRecord =
        reinterpret_cast<RecordConstructor>(bb::oart::GuardedAddress(oart, kRecordConstructor));
    functions.clearSlots =
        reinterpret_cast<ClearSlots>(bb::oart::GuardedAddress(oart, kClearSlots));
    functions.constructImageRecord = reinterpret_cast<ImageRecordConstructor>(
        bb::oart::GuardedAddress(oart, kImageRecordConstructor));
    functions.installCachedImage = reinterpret_cast<InstallCachedImage>(
        bb::oart::GuardedAddress(oart, kInstallCachedImage));
    functions.transferImageSlot = reinterpret_cast<TransferImageSlot>(
        bb::oart::GuardedAddress(oart, kTransferImageSlot));
    functions.buildStretchHolder = reinterpret_cast<BuildStretchHolder>(
        bb::oart::GuardedAddress(oart, kStretchHolderBuilder));
    functions.setCountedSlot =
        reinterpret_cast<SetCountedSlot>(bb::oart::GuardedAddress(oart, kSetCountedSlot));
    functions.constructTransaction = reinterpret_cast<TransactionConstructor>(
        bb::oart::GuardedAddress(oart, kTransactionConstructor));
    functions.destroyRecord =
        reinterpret_cast<Destructor>(bb::oart::GuardedAddress(oart, kRecordDestructor));
    functions.destroyImageRecord =
        reinterpret_cast<Destructor>(bb::oart::GuardedAddress(oart, kImageRecordDestructor));
    functions.destroyTransaction =
        reinterpret_cast<Destructor>(bb::oart::GuardedAddress(oart, kTransactionDestructor));
    functions.destroyHolder =
        reinterpret_cast<Destructor>(bb::oart::GuardedAddress(oart, kHolderDestructor));
    verified = functions;
    verifiedBase = oart;
    return functions;
}

/**
 * Builds the record, commits it through the receiver and destroys everything it
 * built, for one Shape and one already-created cached image.
 *
 * @p cached is borrowed. `+0x8F94C` AddRefs it, so the caller keeps its own
 * reference and may pass the same cached image to any number of Shapes.
 * All temporary Office objects are destroyed before returning, in the reverse
 * order the real handler uses.
 */
/**
 * How many times the private OART apply has been entered this process.
 *
 * This exists so a regression test can *prove* a refusal happened before the
 * dangerous call rather than merely observing that PowerPoint survived. A
 * Connector must leave this untouched. One non-atomic increment on an operation
 * that costs ~190 microseconds is not measurable; STA-only, like everything else
 * here.
 */
unsigned long g_applyEntries = 0;

unsigned long NativeApplyEntryCount() noexcept {
    return g_applyEntries;
}

void ApplyCachedImage(const ApplyFunctions& functions, const FillTarget& target,
                      void* cachedImage, const StageSampler& sample) {
    // Counted at the top: entering at all is what the test is asking about.
    ++g_applyEntries;
    CountedPointer cachedStorage;
    cachedStorage.value = cachedImage;
    CountedPointer* cached = &cachedStorage;
    alignas(16) std::uint8_t recordBuffer[kRecordBufferSize]{};
    alignas(16) std::uint8_t imageRecordBuffer[kImageRecordBufferSize]{};
    alignas(16) std::uint8_t transactionBuffer[kTransactionBufferSize]{};
    static_assert(sizeof(recordBuffer) >= kRecordSize, "record buffer too small");
    static_assert(sizeof(imageRecordBuffer) >= kImageRecordSize, "sub-record buffer too small");
    static_assert(sizeof(transactionBuffer) >= kTransactionSize, "transaction buffer too small");

    if (sample) { sample(L"enter"); }

    void* record = recordBuffer;
    functions.constructRecord(record);
    ConstructedObject recordGuard(record, functions.destroyRecord);
    functions.clearSlots(record);

    // Slot at +0x00: the handler writes the payload then stamps the tag.
    const std::uint32_t picture = kPictureFillKind;
    std::memcpy(recordBuffer + kFillKindPayloadOffset, &picture, sizeof(picture));
    std::uint32_t fillKindTag = MarkSlotSet(bb::oart::LoadDword(record, kFillKindSlotOffset));
    std::memcpy(recordBuffer + kFillKindSlotOffset, &fillKindTag, sizeof(fillKindTag));
    if (sample) { sample(L"record"); }

    void* imageRecord = imageRecordBuffer;
    functions.constructImageRecord(imageRecord);
    ConstructedObject imageRecordGuard(imageRecord, functions.destroyImageRecord);
    if (sample) { sample(L"subrecord"); }

    // +0x8F94C AddRefs the cached image, so our own reference stays ours.
    functions.installCachedImage(imageRecord, cached);
    if (sample) { sample(L"install"); }
    if (bb::oart::LoadPointer(imageRecord, kCachedImageInSubRecordOffset) !=
        reinterpret_cast<std::uintptr_t>(cached->value)) {
        throw bb::Error(E_FAIL, "Cached image did not reach the image sub-record");
    }

    // Copies the sub-record into the record's image slot and AddRefs again.
    functions.transferImageSlot(recordBuffer + kImageSlotOffset, imageRecord);
    if (sample) { sample(L"transfer"); }
    if (bb::oart::LoadPointer(record,
                              kImageSubRecordOffset + kCachedImageInSubRecordOffset) !=
        reinterpret_cast<std::uintptr_t>(cached->value)) {
        throw bb::Error(E_FAIL, "Cached image did not reach the property record");
    }

    // Stretch rather than tile. The handler builds this from sixteen zero bytes;
    // UserTextured differs only by building it another way.
    const std::uint8_t stretchValue[kStretchValueSize]{};
    CountedHolder stretchHolder;
    functions.buildStretchHolder(&stretchHolder, stretchValue);
    HolderGuard stretchGuard(&stretchHolder, functions.destroyHolder);
    functions.setCountedSlot(recordBuffer + kStretchSlotOffset, &stretchHolder);

    // Trailing tagged flag, written exactly as the handler writes it.
    const std::uint8_t trailingPayload = 1;
    std::memcpy(recordBuffer + kTrailingFlagPayloadOffset, &trailingPayload,
                sizeof(trailingPayload));
    std::uint32_t trailingTag =
        MarkSlotSet(bb::oart::LoadDword(record, kTrailingFlagSlotOffset));
    std::memcpy(recordBuffer + kTrailingFlagSlotOffset, &trailingTag, sizeof(trailingTag));
    if (sample) { sample(L"holder"); }

    void* transaction = transactionBuffer;
    functions.constructTransaction(transaction, record, kTransactionFlags,
                                   target.handlerFlag != 0, kPictureFillIdentifier);
    ConstructedObject transactionGuard(transaction, functions.destroyTransaction);
    if (sample) { sample(L"transaction"); }

    // The apply itself: the same receiver vtable slot the handler calls.
    auto vtable = reinterpret_cast<void* const*>(bb::oart::LoadPointer(target.receiver, 0));
    auto applyTransaction = reinterpret_cast<ApplyTransaction>(
        *reinterpret_cast<void* const*>(reinterpret_cast<const std::uint8_t*>(vtable) +
                                        kApplyTransactionSlot));
    applyTransaction(target.receiver, transaction);
    if (sample) { sample(L"apply"); }

    // Destroy in the handler's own reverse order, sampling after each step so an
    // unexplained delta is visible rather than averaged away. Each guard is
    // disarmed by Destroy(), so an exception before this point still unwinds.
    transactionGuard.Destroy();
    imageRecordGuard.Destroy();
    recordGuard.Destroy();
    stretchGuard.Destroy();
    if (sample) { sample(L"released"); }
}

CreatedImage CreateCachedImageFromPixels(const void* pixels, std::uint32_t width,
                                         std::uint32_t height, std::int32_t stride) {
    if (!pixels || width == 0 || height == 0) {
        throw bb::Error(E_INVALIDARG, "Pixel buffer, width and height are required");
    }
    const HMODULE gfx = bb::oart::RequireSupportedModule(L"gfx.dll", "GFX");
    const auto gfxBase = reinterpret_cast<std::uintptr_t>(gfx);
    auto create = reinterpret_cast<CreateCachedImageFromPixelBuffer>(
        reinterpret_cast<void*>(GetProcAddress(gfx, kCreateFromPixelsSymbol)));
    if (!create) {
        throw bb::Error(E_NOTIMPL, "GFX does not export the raw-pixel cached-image creator");
    }

    const Vector2 dpi;
    CountedReference cached;
    CountedReference image;
    create(&cached.storage, &image.storage, pixels, width, height, stride,
           kSurfaceFormatBgra32, &dpi);
    if (!cached.storage.value) {
        throw bb::Error(E_FAIL, "Creator returned no cached image for these pixels");
    }
    if (bb::oart::LoadPointer(cached.storage.value, 0) != gfxBase + kCachedImageVtableRva) {
        cached.storage.value = nullptr;   // unknown layout: leak rather than corrupt
        image.storage.value = nullptr;
        throw bb::Error(E_NOTIMPL, "Cached image vtable does not match the validated layout");
    }
    CreatedImage created{cached.storage.value, image.storage.value};
    cached.storage.value = nullptr;
    image.storage.value = nullptr;
    return created;
}

CreatedImage CreateCachedImageFromBytes(SAFEARRAY* bytes) {
    CountedReference cached;
    CountedReference image;
    const HMODULE gfx = bb::oart::RequireSupportedModule(L"gfx.dll", "GFX");
    const auto gfxBase = reinterpret_cast<std::uintptr_t>(gfx);
    auto createCachedImage = reinterpret_cast<CreateCachedImageFromStream>(
        reinterpret_cast<void*>(GetProcAddress(gfx, kCreateFromStreamSymbol)));
    if (!createCachedImage) {
        throw bb::Error(E_NOTIMPL, "GFX does not export the stream cached-image creator");
    }

    OwnedStream stream = CreateStreamOverBytes(bytes);
    const std::uint8_t uid[kUidSize]{};
    createCachedImage(&cached.storage, &image.storage, stream.value, kStreamCopyInstruction,
                      uid, kCreateFlag);
    if (!cached.storage.value) {
        throw bb::Error(E_FAIL, "Creator returned no cached image for these bytes");
    }
    if (bb::oart::LoadPointer(cached.storage.value, 0) != gfxBase + kCachedImageVtableRva) {
        cached.storage.value = nullptr;   // unknown layout: leak rather than corrupt
        image.storage.value = nullptr;
        throw bb::Error(E_NOTIMPL, "Cached image vtable does not match the validated layout");
    }
    // Ownership transfers to the caller: disarm the scope guards.
    CreatedImage created{cached.storage.value, image.storage.value};
    cached.storage.value = nullptr;
    image.storage.value = nullptr;
    return created;
}

} // namespace bb::oart

