#pragma once
/**
 * @file native_apply.hpp
 * The native picture-fill apply, shared by the research experiments and the
 * texture store.
 *
 * One implementation, one set of guards. Both callers go through
 * ResolveApplyFunctions - which byte-verifies every private entry point before
 * returning any of them - and then ApplyCachedImage.
 *
 * Threading: STA only. The OART FillFormat's reference count is a non-atomic
 * increment, so nothing here may be touched off the owning thread.
 *
 * Ownership: ApplyCachedImage borrows the cached image. Office AddRefs it while
 * the fill is in place, so the same cached image may be applied to any number of
 * Shapes and the caller keeps its own reference throughout.
 */

#include "oart_layout.hpp"

#include <functional>

namespace bb::oart {

// MinGW's x86-64 target already uses the Microsoft convention; __stdcall is
// written out to document intent at each call site.
using RecordConstructor = void(__stdcall*)(void* record);
using ClearSlots = void(__stdcall*)(void* record);
using ImageRecordConstructor = void(__stdcall*)(void* imageRecord);
using InstallCachedImage = void(__stdcall*)(void* imageRecord, void* countedCachedImage);
using TransferImageSlot = void*(__stdcall*)(void* destinationSlot, void* imageRecord);
using TransactionConstructor = void*(__stdcall*)(void* transaction, const void* record,
                                                 std::uint32_t flags, bool flag,
                                                 std::uint32_t identifier);
using Destructor = void(__stdcall*)(void* object);
using BuildStretchHolder = void(__stdcall*)(void* holder, const void* sixteenBytes);
using SetCountedSlot = void*(__stdcall*)(void* slot, void* holder);

/**
 * Every private entry point the apply needs, resolved and byte-verified.
 * Obtained only from ResolveApplyFunctions, which throws if any signature has
 * moved, so holding one of these is evidence the build was checked.
 */
struct ApplyFunctions {
    RecordConstructor constructRecord = nullptr;
    ClearSlots clearSlots = nullptr;
    ImageRecordConstructor constructImageRecord = nullptr;
    InstallCachedImage installCachedImage = nullptr;
    TransferImageSlot transferImageSlot = nullptr;
    BuildStretchHolder buildStretchHolder = nullptr;
    SetCountedSlot setCountedSlot = nullptr;
    TransactionConstructor constructTransaction = nullptr;
    Destructor destroyRecord = nullptr;
    Destructor destroyImageRecord = nullptr;
    Destructor destroyTransaction = nullptr;
    Destructor destroyHolder = nullptr;
};

/// Byte-verifies every entry point, then returns them. Throws bb::Error if any
/// signature no longer matches the bytes its ABI was derived from.
ApplyFunctions ResolveApplyFunctions(std::uintptr_t oartBase);

/// Optional per-stage hook, used by the experiments to sample reference counts.
using StageSampler = std::function<void(const wchar_t*)>;

/**
 * Builds the property record, commits it through @p target's receiver, and
 * destroys every temporary in the reverse order the real handler uses.
 *
 * @p cachedImage is borrowed; the caller keeps its own reference.
 */
void ApplyCachedImage(const ApplyFunctions& functions, const FillTarget& target,
                      void* cachedImage, const StageSampler& sample = {});

/// One cached image and its companion image, each carrying one owned reference.
struct CreatedImage {
    void* cached = nullptr;
    void* image = nullptr;
};

/**
 * Decodes @p bytes into a GFX cached image through the exported
 * GEL::ICachedImage::Create, with no file on disk.
 *
 * Both returned pointers carry one reference each and belong to the caller.
 * Release them with ReleaseIntrusive.
 */
CreatedImage CreateCachedImageFromBytes(SAFEARRAY* bytes);

} // namespace bb::oart
