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

#include <windows.h>
#include <oleauto.h>

#include <string>

/**
 * True when the native backend can actually run in this process: PowerPoint
 * host, all three Office modules at the validated build, the exported GFX
 * creator present, and every private entry point's signature bytes intact.
 * False on any other host, which is what the capability string must report.
 */
bool nativeTextureBackendAvailable() noexcept;

// Whether a Shape may take the native path is decided by
// bb::office::RequireNativePictureFillTarget in shape_policy.hpp. It is not
// re-declared here: one authority, asked by every entry point.


long nativeTextureLoad(SAFEARRAY* bytes);

/**
 * Loads a texture from a raw 32-bit BGRA buffer instead of an encoded image.
 * Same handles, same lifetime rules; only the decode is skipped.
 */
long nativeTextureLoadPixels(const void* pixels, unsigned long width,
                             unsigned long height, long stride);

void nativeTextureApply(IDispatch* fill, long handle);
void nativeTextureRelease(long handle);
void nativeTextureClear() noexcept;
bool nativeTextureOwnsHandle(long handle);
long nativeTextureCount();

/// Diagnostics: handle count, creations and reference counts, as a `key=value;`
/// string. Used by the Office regression harnesses, not by the shipping path.
std::wstring nativeTextureReport(long handle);
