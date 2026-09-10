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
 * Which of the receiver's entry points the finished transaction goes through.
 *
 * The two are not alternatives in the sense of being interchangeable choices:
 * `Combined` is what Office itself calls, and `Split` is that same call taken
 * apart into the two steps it makes internally. Reading the handler on build
 * 16.0.14334.20848, receiver slot +0x78 is a thunk to slot +0x50, and slot +0x50
 * does exactly this:
 *
 *     stamp the transaction's +0x14 word
 *     if a feature flag is off:
 *         slot +0x60 (receiver, transaction, &produced)   <- computes the change
 *         slot +0x58 (receiver, produced)                 <- commits it
 *         produced->vtable[+0xA8](produced, 1)            <- releases it
 *
 * `Split` exists to find out which of those two steps the cost is in. Measured
 * over 2000 rounds: computing costs 5 microseconds and committing costs 144 -
 * and a change that is computed and never committed does nothing at all, which
 * is how the naming above was settled. It is research only: it reproduces a code
 * path rather than calling it, so it can diverge from what Office does if that
 * feature flag is on, and the way to know is to compare the document and the
 * timings against `Combined`.
 */
enum class ApplyRoute {
    /// The receiver's own entry point. What production always uses.
    Combined,
    /// The same work, driven step by step. Research only; never shipped behind
    /// an API a caller can reach.
    Split,
    /**
     * Computes the change and drops it without committing it.
     *
     * Kept because the answer is worth keeping: it costs 0.013 ms against 0.166,
     * and it **does nothing**. A clean Shape put through it stays at Fill.Type 1
     * and renders byte-identically, measured in tools/test_change_only.ps1. So
     * the 144 microseconds are not bookkeeping wrapped around a cheap edit that
     * could be skipped - they are the edit.
     *
     * Research only, and kept so that negative result stays reproducible.
     */
    ChangeOnly,
};

/**
 * Builds the property record, commits it through @p target's receiver, and
 * destroys every temporary in the reverse order the real handler uses.
 *
 * @p cachedImage is borrowed; the caller keeps its own reference.
 *
 * With `ApplyRoute::Split` the sampler additionally sees `change` and `record`
 * between `transaction` and `apply`, so the two halves can be timed apart.
 */
void ApplyCachedImage(const ApplyFunctions& functions,
                      const FillTarget& target,
                      void* cachedImage,
                      const StageSampler& sample = {},
                      ApplyRoute route = ApplyRoute::Combined);

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
