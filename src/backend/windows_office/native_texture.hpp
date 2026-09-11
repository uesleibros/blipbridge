#pragma once
/**
 * @file native_texture.hpp
 * The reusable texture store the Windows backend is built on.
 *
 * A texture is one decoded image that any number of Shapes can be filled with.
 * It owns one intrusive reference to a GFX cached image and is independent of
 * any document; the ownership rules are in native_texture.cpp and the validation
 * matrix in docs/native_texture.md.
 *
 * This is the seam between `src/backend/windows_office_backend.cpp` and the
 * reverse-engineered implementation. It exists so that production code does not
 * have to include the research header to reach the texture store - the research
 * harnesses include *this*, not the other way round.
 *
 * Threading: STA only, like everything else that touches these Office objects.
 * Handles never recycle, and live in a range disjoint from the donor fallback's
 * so one ApplyTexture can serve both.
 */

// MinGW requires Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

#include <cstdint>
#include <memory>
#include <string>

/**
 * True when the native backend can actually run in this process: PowerPoint
 * host, all three Office modules at the validated build, the exported GFX
 * creator present, and every private entry point's signature bytes intact.
 * False on any other host, which is what the capability string must report.
 */
bool nativeTextureBackendAvailable() noexcept;

namespace bb::office {

/**
 * One decoded image and the GFX references it owns. Defined in
 * native_texture.cpp; callers only ever hold it through TextureRef.
 */
class NativeTexture;

/**
 * Shared ownership of a decoded image.
 *
 * This is the fix for a real defect. The texture store used to own each image
 * outright and hand callers a `long` handle into its map, and the picture cache
 * stored one of those handles - so `ClearTextures()`, or releasing that
 * particular handle, destroyed the image the picture cache was still pointing
 * at, and the next `UserPicture2` failed with "handle ... is not valid (it was
 * released)".
 *
 * The public handle is now an external *token* that references a resource, not
 * the resource itself. Two independent owners can hold the same image:
 *
 *     public handle   -> TextureRef
 *     picture cache   -> TextureRef
 *
 * Dropping either leaves the other working. Released handles still never come
 * back and are never recycled; only the reference behind them goes away.
 */
using TextureRef = std::shared_ptr<NativeTexture>;

/// Decodes bytes into a new image. The caller owns the only reference so far.
TextureRef CreateTextureFromBytes(SAFEARRAY* bytes);

/// Same, from a raw BGRA32 buffer.
TextureRef CreateTextureFromPixels(const void* pixels, unsigned long width,
                                   unsigned long height, long stride);

/**
 * Gives @p texture a public handle, adding one owner.
 *
 * Handles are never recycled, so a handle released later stays permanently
 * stale even though the image behind it may still be alive in the picture
 * cache.
 */
long RegisterTexture(TextureRef texture);

/// Applies an image the caller already holds, with no handle lookup at all.
void ApplyTextureRef(IDispatch* fill, const TextureRef& texture);

/**
 * A process-unique id for one decoded image.
 *
 * The picture cache remembers which image a Shape last received, and it must be
 * able to compare that across cache eviction without keeping the image alive or
 * dereferencing something freed. An id answers exactly that and nothing else.
 */
std::uint64_t TextureIdOf(const TextureRef& texture);

/**
 * Adds @p times to @p texture's apply count.
 *
 * The count exists so a test can show creations staying far below applies -
 * that is what proves the image is being reused rather than rebuilt. A range
 * apply is one private call that fills many Shapes, so it counts one per member
 * and the diagnostic keeps meaning what it meant.
 */
void CountTextureApplies(const TextureRef& texture, unsigned long times);

/**
 * Resolves a public handle to the image it references.
 *
 * Throws bb::Error if the handle is unknown or was released. Used by the stage
 * profiler, which needs the lookup and the apply as separately timed steps
 * rather than as one call.
 */
TextureRef LookupTexture(long handle);

/**
 * The GFX cached image inside @p texture, borrowed.
 *
 * Exists so a profiler can hand the image to ApplyCachedImage directly and time
 * the guard chain separately from the apply. Not an ownership transfer: the
 * TextureRef still owns it.
 */
void* CachedImageOf(const TextureRef& texture);

} // namespace bb::office

// Whether a Shape may take the native path is decided by
// bb::office::RequireNativePictureFillTarget in shape_policy.hpp. It is not
// re-declared here: one authority, asked by every entry point.

long nativeTextureLoad(SAFEARRAY* bytes);

/**
 * Loads a texture from a raw 32-bit BGRA buffer instead of an encoded image.
 * Same handles, same lifetime rules; only the decode is skipped.
 */
long nativeTextureLoadPixels(const void* pixels,
                             unsigned long width,
                             unsigned long height,
                             long stride);

void nativeTextureApply(IDispatch* fill, long handle);

/**
 * Applies @p handle to the Shape @p shape, fetching its Fill.
 *
 * Unlike the Fill-only overload this keeps the shared per-Shape record truthful,
 * so a later ApplyTextureIfChanged or BB_ApplyPicture cannot skip on the
 * strength of an image this call replaced. @p shapeType is `Shape.Type`, already
 * read by the caller's eligibility check.
 */
void nativeTextureApplyToShape(IDispatch* shape, long handle, long shapeType);

/**
 * Applies @p handle to @p shape only if it is not already carrying that image.
 *
 * @p shapeType is `Shape.Type`, passed in because the caller has already read it
 * while deciding the Shape was eligible. @p skipped receives whether the Office
 * edit was avoided, and is written even when this throws.
 *
 * The identity compared is the process-unique image id, never the public handle:
 * a handle can be released while another owner keeps the same image alive.
 *
 * Every check the ordinary apply makes still runs. This only removes the Office
 * document edit, and only when the Shape demonstrably does not need one.
 */
void nativeTextureApplyIfChanged(IDispatch* shape, long handle, long shapeType, bool* skipped);
void nativeTextureRelease(long handle);
void nativeTextureClear() noexcept;
bool nativeTextureOwnsHandle(long handle);
long nativeTextureCount();

/// Diagnostics: handle count, creations and reference counts, as a `key=value;`
/// string. Used by the Office regression harnesses, not by the shipping path.
std::wstring nativeTextureReport(long handle);
