#pragma once
// MinGW requires Windows base types before the Automation declarations.
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
std::wstring
benchmarkNativeTexture(IDispatch* slide, const std::wstring& imagePath, long iterations);
/**
 * Compares one BB_ApplyTextureBatch call against the same number of individual
 * BB_ApplyTexture calls, through the real C ABI, in process. Creates and deletes
 * its own Shapes.
 */
std::wstring benchmarkTextureBatch(IDispatch* slide, long shapeCount, long iterations);

/**
 * Research: the six scenarios that decide whether BB_ApplyTextureIfChanged is
 * worth having - repeated applies, a first apply, repeated skips, alternating
 * images, Shape deletion and recreation, and many Shapes sharing one texture.
 * Asserts the skip decision each leg expected. See skip_benchmark.cpp.
 */
std::wstring benchmarkApplySkip(IDispatch* slide, long shapeCount, long iterations);

/**
 * Research: what the private apply's cost is made of, by varying one thing at a
 * time around it - displayed slide, another slide, hidden, off-slide, minimised
 * window. Answers whether the 76% is rendering or document bookkeeping, which
 * decides whether batching is worth building. See apply_cost_factors.cpp.
 */
std::wstring measureApplyCostFactors(IDispatch* presentation, long iterations);

/**
 * Research: step-by-step timing of ResolveFillTarget, the second largest cost in
 * an apply and the largest one that is ours. See resolve_profiler.cpp.
 */
std::wstring profileResolveStages(IDispatch* fill, long iterations);

/**
 * Research: the private apply run both through the receiver's own entry point
 * and through the two steps it takes internally, to find out which half costs
 * the 0.15 ms - which is what decides whether a multi-Shape batch could exist
 * at all. See transaction_split.cpp.
 */
std::wstring splitTransactionApply(IDispatch* shape, long handle, long iterations);

/**
 * Research: one apply that performs the change and does not record it, so a
 * harness can find out what the recording was buying. See transaction_split.cpp.
 */
std::wstring applyChangeOnly(IDispatch* shape, long handle);

/**
 * Research: applies a cached image straight to a FillFormat, so a ShapeRange's
 * fill can be asked whether one private apply fills every Shape in it. Skips the
 * semantic gate by construction; see transaction_split.cpp.
 */
std::wstring applyCachedImageToFill(IDispatch* fill, long handle, long iterations);

/**
 * Benchmarks the three public ways to fill N Shapes - N x BB_ApplyTexture,
 * BB_ApplyTextureBatch and BB_ApplyTextureRange - over one ShapeRange, all
 * through the C ABI in process. The API itself is production; only this timing
 * harness is research. See range_apply.cpp.
 */
std::wstring applyTextureToRange(IDispatch* range, long handle, long iterations);

/**
 * Research: builds a cached image from raw pixels through the exported GFX
 * raw-pixel creator and applies it. The surface-format value is a parameter
 * because ARC::SurfaceFormat has no symbols; the harness probes it.
 */
std::wstring pixelTextureExperiment(
    IDispatch* fill, SAFEARRAY* pixels, long width, long height, long stride, long surfaceFormat);

/**
 * Research: reports how far a Shape gets along the validated fill chain, naming
 * the step it stops at rather than throwing at the first failure. Read-only.
 * Feeds tools/test_shape_compatibility.ps1; see shape_compatibility.cpp.
 */
std::wstring probeShapeCompatibility(IDispatch* shape);

/**
 * Research: the semantic verdict for a Shape, plus the running count of entries
 * into the private OART apply.
 *
 * The count is what lets a regression test prove a refusal happened *before* the
 * dangerous call, rather than merely observing that PowerPoint survived it.
 */
std::wstring probeShapePolicy(IDispatch* shape);

/**
 * Research: stage-by-stage timing of the production ApplyTexture path.
 *
 * Walks exactly what BB_ApplyTexture walks, timing each step separately, so the
 * ~0.186 ms can be attributed rather than guessed at. See apply_profiler.cpp.
 */
std::wstring profileApplyStages(IDispatch* shape, long handle, long iterations);

/**
 * Research: applies a texture with the Shape-type allowlist bypassed, so the
 * matrix can find out which classes are genuinely safe rather than merely
 * structurally identical. Connectors stay refused - that case is settled.
 * Never reachable through the C ABI. See shape_compatibility.cpp.
 */
std::wstring applyTextureUnrestricted(IDispatch* shape, long handle);

/**
 * Research drivers for BB_ApplyPicture and its cache controls, forwarding to the
 * real C ABI in process so the harness asserts shipping behaviour rather than a
 * re-implementation. See picture_cache_driver.cpp.
 */
std::wstring applyPictureThroughAbi(IDispatch* shape, const std::wstring& path);
std::wstring applyTextureIfChangedThroughAbi(IDispatch* shape, long handle);
std::wstring applyTextureRangeThroughAbi(IDispatch* range, long handle);
std::wstring capabilitiesThroughAbi();
std::wstring invalidateShapeThroughAbi(IDispatch* shape);
std::wstring clearPictureCacheThroughAbi();
std::wstring pictureCacheStatsThroughAbi();

/**
 * Research: stage-by-stage attribution of the picture-fill path, across four
 * legs (reuse, new-same-bytes, new-distinct-content, and a pre-created pool).
 * Answers where the difference between applying an existing texture and
 * creating new content actually goes. See stage_profiler.cpp.
 */
std::wstring profileFillStages(IDispatch* slide, const std::wstring& imagePath, long iterations);

/** Compares encoded-image loading against raw-pixel loading at one size. */
std::wstring
benchmarkPixelLoad(const std::wstring& imagePath, long width, long height, long iterations);
