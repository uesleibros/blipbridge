#pragma once
/**
 * @file apply_skip_cache.hpp
 * Remembering which image a Shape already carries, so a redundant apply can be
 * skipped without entering Office at all.
 *
 * ## Why this exists
 *
 * A cached `ApplyTexture` costs about 0.19 ms, and the profiler attributes 76%
 * of that to one private OART call - the document edit, its undo entry and the
 * invalidation that follows. Office does not short-circuit an apply that changes
 * nothing: applying the same image twice costs the same as applying two
 * different ones, measured. So the only way to avoid that cost is to not make
 * the call, and the only safe way to do *that* is to know the Shape already has
 * the image.
 *
 * ## Identity, and why not the public handle
 *
 * The remembered image is a **process-unique image id**, never the caller's
 * texture handle. A handle can be released while the picture cache still holds
 * the same image alive, and a later handle for that same image would be a
 * different number - so handle equality answers a different question than "is
 * this the same picture". `TextureIdOf` answers the right one.
 *
 * ## One record, shared
 *
 * Both skipping paths - `BB_ApplyPicture`'s cache and `ApplyTextureIfChanged` -
 * use *this* map, and every path that changes a fill updates it. Two maps over
 * the same Shapes would be worse than none: one could say a Shape carries image
 * X while the other had just put image Y on it, and the next skip would honour
 * the stale one. A wrong picture, silently, is the single failure this must
 * never produce, so there is exactly one record and everything that writes a
 * fill writes to it.
 *
 * ## What is never cached
 *
 * No OART receiver, no control block, no `FillFormat`, no Shape pointer. Only
 * value identities. A deleted Shape leaves a structurally plausible pointer
 * chain that passes every check, which is exactly why the receiver is re-resolved
 * on every real apply and why nothing here holds one.
 */

// MinGW requires the Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

#include <cstddef>
#include <cstdint>

#include "shape_identity.hpp"

namespace bb::office {

/// What ApplyTextureIfChanged decided to do, so a caller can tell them apart.
enum class SkipDecision {
    /// The apply ran: the Shape did not already carry this image.
    Applied,
    /// The Shape already carried it, verified, and nothing was touched.
    Skipped,
};

/**
 * Records that @p shape now carries the image @p textureId.
 *
 * Does nothing for an unkeyable Shape - a group, or anything inside one.
 */
void RememberApplied(const ShapeKey& key, std::uint64_t textureId) noexcept;

/**
 * True when @p shape already carries @p textureId and still looks like it.
 *
 * Before answering yes it re-reads `Fill.Type` and requires a picture fill, so a
 * Shape whose fill was cleared or replaced with a colour is re-applied rather
 * than wrongly skipped. What that check cannot see is a fill replaced with a
 * *different* picture by something else; detecting that would cost more than the
 * apply it saves, so `BB_InvalidateShape` is the documented remedy.
 */
bool AlreadyCarries(const ShapeKey& key, std::uint64_t textureId, IDispatch* shape) noexcept;

/// Forgets one Shape. Forgetting an unknown Shape is not an error.
void ForgetApplied(const ShapeKey& key) noexcept;

/// Forgets every Shape. Called when texture ownership changes wholesale.
void ForgetAllApplied() noexcept;

/**
 * Forgets every Shape remembered as carrying @p textureId.
 *
 * Used when an image is released, so a later image cannot inherit a stale claim.
 */
void ForgetTexture(std::uint64_t textureId) noexcept;

/**
 * True when any Shape at all is remembered.
 *
 * Lets a path that does not itself skip - the ordinary apply - decide cheaply
 * whether it needs to keep the record truthful. A caller who never asks for a
 * skip pays one test rather than a Shape key.
 */
bool AnyRemembered() noexcept;

/// How many Shapes are remembered, and how many applies have been skipped.
struct SkipStats {
    std::size_t shapes = 0;
    std::uint64_t skipped = 0;
};

SkipStats ApplySkipStats() noexcept;

} // namespace bb::office
