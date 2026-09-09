#pragma once
#include <windows.h>
#include <oleauto.h>
#include <string>

/** Research entry points: synchronous PowerPoint STA only, never normal backends. */
int runExperiment(int argc, wchar_t** argv);
/** Temporarily instruments file APIs and restores patches before returning. */
HRESULT traceUserPicture(IDispatch* fill, const std::wstring& root);
HRESULT focusedBenchmarks(IDispatch* application, const std::wstring& root);
/**
 * Delivers borrowed bytes through a temporary file-API adapter. Office still
 * creates a Content.MSO image file; this does not satisfy the file-free goal.
 */
HRESULT memoryFillExperiment(IDispatch* shape, SAFEARRAY* bytes, const std::wstring& root);
HRESULT stressExperiment(IDispatch* application, const std::wstring& root);
HRESULT traceCachedApply(IDispatch* donor, IDispatch* target, const std::wstring& root);
/**
 * Read-only, version- and vtable-guarded walk from a PowerPoint FillFormat to
 * the OART receiver behind it. Calls no private Office function and retains
 * nothing.
 *
 * Throws bb::Error naming the check that failed, so an unsupported build or an
 * unexpected object layout is diagnosable rather than a bare E_NOTIMPL. The
 * caller runs inside Engine::Invoke's handler, which converts it to EXCEPINFO.
 */
std::wstring inspectFillReceiver(IDispatch* fill);
/**
 * Decodes bytes already in memory into an Office cached image through the
 * exported GFX stream creator, then releases it. Touches no document state.
 * Throws bb::Error naming the check that failed.
 */
std::wstring loadCachedImageExperiment(SAFEARRAY* bytes);
/**
 * Applies bytes as a picture fill on the Shape behind a FillFormat, natively:
 * memory IStream to cached image to OART record to transaction to the receiver.
 * No UserPicture, no donor Shape and no source file. Retains nothing.
 */
std::wstring nativeApplyExperiment(IDispatch* fill, SAFEARRAY* bytes);
/**
 * Decodes the bytes once and applies that single cached image to every supplied
 * FillFormat, re-resolving each Shape's receiver immediately before its apply.
 * Reports the one cached-image address and the per-Shape reference timeline.
 */
std::wstring nativeApplyReuseExperiment(SAFEARRAY* fills, SAFEARRAY* bytes);

/**
 * Reusable texture handles over GFX cached images. A texture owns one decoded
 * image and is independent of any document; see native_texture.cpp for the
 * ownership rules and docs/native_texture.md for the validation matrix.
 *
 * STA only. Handles never recycle, and live in a range disjoint from the donor
 * fallback's so one ApplyTexture can serve both.
 */
/**
 * True when the native backend can actually run in this process: PowerPoint
 * host, all three Office modules at the validated build, the exported GFX
 * creator present, and every private entry point's signature bytes intact.
 * False on any other host, which is what the capability string must report.
 */
bool nativeTextureBackendAvailable() noexcept;
long nativeTextureLoad(SAFEARRAY* bytes);
void nativeTextureApply(IDispatch* fill, long handle);
void nativeTextureRelease(long handle);
void nativeTextureClear() noexcept;
bool nativeTextureOwnsHandle(long handle);
long nativeTextureCount();
std::wstring nativeTextureReport(long handle);

/**
 * In-process comparison of Fill.UserPicture against LoadTexture once plus
 * repeated ApplyTexture, on one Shape of the supplied Slide. Reports the load
 * cost separately from the hot-path apply cost. Leaves the document as found.
 */
std::wstring benchmarkNativeTexture(IDispatch* slide, const std::wstring& imagePath,
                                    long iterations);
/**
 * Compares one BB_ApplyTextureBatch call against the same number of individual
 * BB_ApplyTexture calls, through the real C ABI, in process. Creates and deletes
 * its own Shapes.
 */
std::wstring benchmarkTextureBatch(IDispatch* slide, long shapeCount, long iterations);
