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
using TransactionConstructor = void*(__stdcall*)(void* transaction,
                                                 const void* record,
                                                 std::uint32_t flags,
                                                 bool flag,
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

/**
 * Optional per-stage hook, used by the experiments to sample reference counts
 * and, in the stage profiler, wall-clock time.
 *
 * It is called after each step of the apply, in this order:
 *
 *   enter, record, subrecord, install, transfer, holder, transaction, apply,
 *   released
 *
 * `enter` is a baseline taken before any work, so the cost of a stage is the
 * difference between its sample and the previous one. When no sampler is
 * supplied - which is every production call - the only cost is one null test
 * per stage.
 */
using StageSampler = std::function<void(const wchar_t*)>;

/**
 * Builds the property record, commits it through @p target's receiver, and
 * destroys every temporary in the reverse order the real handler uses.
 *
 * @p cachedImage is borrowed; the caller keeps its own reference.
 */
void ApplyCachedImage(const ApplyFunctions& functions,
                      const FillTarget& target,
                      void* cachedImage,
                      const StageSampler& sample = {});

/**
 * How many times ApplyCachedImage has been entered this process.
 *
 * Diagnostic, for the regression tests that must show a Shape was refused
 * *before* the private OART apply, not merely that the host survived it.
 */
unsigned long NativeApplyEntryCount() noexcept;

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

/**
 * Builds a cached image from a raw 32-bit BGRA buffer, through the exported
 * GEL::ICachedImage::Create overload that takes a pixel pointer.
 *
 * Only BGRA32 is offered. The surface-format argument that overload takes was
 * probed across values 0 to 24 and every one produced an identical, correct
 * BGRA render, so the value carries no information this code could honestly act
 * on; a fixed value is passed and other layouts are the caller's to convert.
 * See docs/pixel_textures.md.
 *
 * @p pixels is borrowed for the duration of the call. Both returned pointers
 * carry one reference each and belong to the caller.
 */
CreatedImage CreateCachedImageFromPixels(const void* pixels,
                                         std::uint32_t width,
                                         std::uint32_t height,
                                         std::int32_t stride);

} // namespace bb::oart
