#pragma once
#include <windows.h>
#include <oleauto.h>
#include <string>

// The texture store is production code and lives with the backend. The research
// harnesses below drive it; it does not know they exist.
#include "../src/backend/windows_office/native_texture.hpp"

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

/**
 * Research: builds a cached image from raw pixels through the exported GFX
 * raw-pixel creator and applies it. The surface-format value is a parameter
 * because ARC::SurfaceFormat has no symbols; the harness probes it.
 */
std::wstring pixelTextureExperiment(IDispatch* fill, SAFEARRAY* pixels, long width,
                                    long height, long stride, long surfaceFormat);

/**
 * Research: reports how far a Shape gets along the validated fill chain, naming
 * the step it stops at rather than throwing at the first failure. Read-only.
 * Feeds tools/test_shape_compatibility.ps1; see shape_compatibility.cpp.
 */
std::wstring probeShapeCompatibility(IDispatch* shape);

/**
 * Research: applies a texture with the Shape-type allowlist bypassed, so the
 * matrix can find out which classes are genuinely safe rather than merely
 * structurally identical. Connectors stay refused - that case is settled.
 * Never reachable through the C ABI. See shape_compatibility.cpp.
 */
std::wstring applyTextureUnrestricted(IDispatch* shape, long handle);

/**
 * Research: stage-by-stage attribution of the picture-fill path, across four
 * legs (reuse, new-same-bytes, new-distinct-content, and a pre-created pool).
 * Answers where the difference between applying an existing texture and
 * creating new content actually goes. See stage_profiler.cpp.
 */
std::wstring profileFillStages(IDispatch* slide, const std::wstring& imagePath,
                               long iterations);

/** Compares encoded-image loading against raw-pixel loading at one size. */
std::wstring benchmarkPixelLoad(const std::wstring& imagePath, long width, long height,
                                long iterations);
