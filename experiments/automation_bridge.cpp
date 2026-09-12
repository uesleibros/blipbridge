#include "../src/com/engine.hpp"
#include "experiment_api.hpp"

#if !defined(BB_HAS_NATIVE_BACKEND)
#include "../src/backend/portable_office/portable_texture.hpp"
#endif

#include <string>

namespace bb {
namespace {
/// Argument counts are part of the research ABI; keep them in one place.
UINT ExpectedResearchArgumentCount(DispatchId id) {
    switch (id) {
    case DispatchId::ClearPictureCache:
    case DispatchId::PictureCacheStats:
    case DispatchId::AbiCapabilities:
    case DispatchId::ClearImagesAbi:
    case DispatchId::AbiCounts:
        return 0;
    case DispatchId::RunBenchmarks:
    case DispatchId::ProbeShapeCompatibility:
    case DispatchId::ProbeShapePolicy:
    case DispatchId::InvalidateShape:
    case DispatchId::InspectFillReceiver:
    case DispatchId::LoadCachedImageExperiment:
    case DispatchId::InspectTexture:
    case DispatchId::ReleaseImageAbi:
    case DispatchId::LoadTextureBytesAbi:
    case DispatchId::ImageSizeAbi:
    case DispatchId::CreateTextureFromImageAbi:
    case DispatchId::AbiLifecycle:
        return 1;
    case DispatchId::PixelTextureExperiment:
        return 6;
    case DispatchId::BenchmarkPixelLoad:
    case DispatchId::LoadTexturePixelsAbi:
    case DispatchId::LoadImagePixelsAbi:
        return 4;
    case DispatchId::MemoryFillExperiment:
    case DispatchId::TraceCachedApply:
    case DispatchId::BenchmarkNativeTexture:
    case DispatchId::BenchmarkTextureBatch:
    case DispatchId::BenchmarkApplySkip:
    case DispatchId::SplitTransactionApply:
    case DispatchId::ApplyCachedImageToFill:
    case DispatchId::ApplyTextureToRange:
    case DispatchId::WarpApplyQuad:
    case DispatchId::ApplyImageQuadAbi:
    case DispatchId::ProfileFillStages:
    case DispatchId::ProfileApplyStages:
    case DispatchId::WarpImageQuadAbi:
        return 3;
    default:
        return 2;
    }
}
#if !defined(BB_HAS_NATIVE_BACKEND)
/**
 * The answer for a probe that only exists over the accelerated backend.
 *
 * These methods instrument reverse-engineered Office internals - receiver
 * layouts, transaction splitting, GFX image lifetimes - so there is nothing for
 * them to instrument in a portable build. Keeping the DispatchId and refusing it
 * by name is deliberate: the ids are a stable Automation ABI, and a harness that
 * asks for one gets told what is missing rather than "unknown member", which
 * would read like a typo in the harness.
 */
[[noreturn]] void RequireAcceleratedBackend(const char* probe) {
    throw Error(E_NOTIMPL,
                std::string(probe) +
                    " instruments the accelerated Office backend, which is not compiled in this "
                    "build. The portable backend has no Office internals to inspect");
}
#endif

} // namespace

/**
 * Explicit research-only Automation routing. Lives with the experiments so
 * production texture operations cannot silently acquire instrumentation hooks.
 * Argument Values retain COM objects/SAFEARRAYs for the synchronous call only.
 */
Value Engine::DispatchResearch(DispatchId id, const AutomationArguments& arguments) {
    const UINT expectedCount = ExpectedResearchArgumentCount(id);
    arguments.RequireCount(expectedCount);
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw Error(E_ACCESSDENIED, "Research methods require PowerPoint host");
    }

    if (id == DispatchId::RunBenchmarks) {
        auto root = arguments.At(0).str();
        wchar_t executableName[] = L"bb";
        wchar_t* argv[] = {executableName, root.data()};
        return Value(static_cast<long>(runExperiment(2, argv)));
    }

    // Handled before the target is extracted: these take no arguments at all,
    // and At(0) would throw for them.
    if (id == DispatchId::ClearPictureCache) {
        return Value(clearPictureCacheThroughAbi().c_str());
    }
    if (id == DispatchId::AbiCapabilities) {
        return Value(capabilitiesThroughAbi().c_str());
    }
    if (id == DispatchId::PictureCacheStats) {
        return Value(pictureCacheStatsThroughAbi().c_str());
    }
    if (id == DispatchId::ClearImagesAbi) {
        return Value(clearImagesThroughAbi().c_str());
    }
    if (id == DispatchId::AbiCounts) {
        return Value(abiCountsThroughAbi().c_str());
    }

    auto target = arguments.At(0);
    switch (id) {
    case DispatchId::TraceUserPicture:
        check(traceUserPicture(target.obj(), arguments.At(1).str()), "TraceUserPicture");
        break;
    case DispatchId::RunFocusedBenchmarks:
        check(focusedBenchmarks(target.obj(), arguments.At(1).str()), "Focused benchmarks");
        break;
    case DispatchId::MemoryFillExperiment: {
        auto bytes = arguments.At(1);
        if (bytes.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        check(memoryFillExperiment(target.obj(), bytes.v.parray, arguments.At(2).str()),
              "Memory fill experiment");
        break;
    }
    case DispatchId::RunStress:
        check(stressExperiment(target.obj(), arguments.At(1).str()), "Stress experiment");
        break;
    case DispatchId::InspectTexture:
        // Both stores answer this, in the same key=value shape, so a harness
        // asking what a handle holds gets an answer on either backend.
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(nativeTextureReport(target.integer()).c_str());
#else
        return Value(
            bb::portable::DescribeTextures(static_cast<std::uint64_t>(target.integer())).c_str());
#endif
    case DispatchId::BenchmarkNativeTexture:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(
            benchmarkNativeTexture(target.obj(), arguments.At(1).str(), arguments.At(2).integer())
                .c_str());
#else
        RequireAcceleratedBackend("BenchmarkNativeTexture");
#endif
    case DispatchId::BenchmarkTextureBatch:
        return Value(benchmarkTextureBatch(
                         target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                         .c_str());
    case DispatchId::LoadImageAbi:
        return Value(loadImageThroughAbi(target.str(), arguments.At(1).str()).c_str());
    case DispatchId::LoadTextureScaledAbi:
        return Value(loadTextureScaledThroughAbi(target.str(), arguments.At(1).str()).c_str());
    case DispatchId::ReleaseImageAbi:
        return Value(releaseImageThroughAbi(target.integer()).c_str());
    case DispatchId::ApplyImageQuadAbi: {
        auto points = arguments.At(2);
        if (points.v.vt != (VT_ARRAY | VT_R8) && points.v.vt != (VT_ARRAY | VT_VARIANT)) {
            throw Error(E_INVALIDARG, "Expected an array of eight quad coordinates");
        }
        return Value(
            applyImageQuadThroughAbi(target.obj(), arguments.At(1).integer(), points.v.parray)
                .c_str());
    }
    case DispatchId::ProbeDynamicTexture:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(probeDynamicTexture(target.obj(), arguments.At(1).integer()).c_str());
#else
        RequireAcceleratedBackend("ProbeDynamicTexture");
#endif
    case DispatchId::WarpApplyQuad:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        auto points = arguments.At(2);
        if (points.v.vt != (VT_ARRAY | VT_R8) && points.v.vt != (VT_ARRAY | VT_VARIANT)) {
            throw Error(E_INVALIDARG, "Expected an array of eight quad coordinates");
        }
        return Value(
            warpApplyQuadFromFile(target.obj(), arguments.At(1).str(), points.v.parray).c_str());
    }
#else
        RequireAcceleratedBackend("WarpApplyQuad");
#endif
    case DispatchId::ApplyTextureRange:
        return Value(applyTextureRangeThroughAbi(target.obj(), arguments.At(1).integer()).c_str());
    case DispatchId::ApplyTextureToRange:
        return Value(
            applyTextureToRange(target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                .c_str());
    case DispatchId::ApplyCachedImageToFill:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(applyCachedImageToFill(
                         target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                         .c_str());
#else
        RequireAcceleratedBackend("ApplyCachedImageToFill");
#endif
    case DispatchId::ApplyChangeOnly:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(applyChangeOnly(target.obj(), arguments.At(1).integer()).c_str());
#else
        RequireAcceleratedBackend("ApplyChangeOnly");
#endif
    case DispatchId::SplitTransactionApply:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(splitTransactionApply(
                         target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                         .c_str());
#else
        RequireAcceleratedBackend("SplitTransactionApply");
#endif
    case DispatchId::ApplyTextureIfChanged:
        return Value(
            applyTextureIfChangedThroughAbi(target.obj(), arguments.At(1).integer()).c_str());
    case DispatchId::ProfileResolveStages:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(profileResolveStages(target.obj(), arguments.At(1).integer()).c_str());
#else
        RequireAcceleratedBackend("ProfileResolveStages");
#endif
    case DispatchId::MeasureApplyCostFactors:
        return Value(measureApplyCostFactors(target.obj(), arguments.At(1).integer()).c_str());
    case DispatchId::BenchmarkApplySkip:
        return Value(
            benchmarkApplySkip(target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                .c_str());
    case DispatchId::ProfileApplyStages:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(
            profileApplyStages(target.obj(), arguments.At(1).integer(), arguments.At(2).integer())
                .c_str());
#else
        RequireAcceleratedBackend("ProfileApplyStages");
#endif
    case DispatchId::ProfileFillStages:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(
            profileFillStages(target.obj(), arguments.At(1).str(), arguments.At(2).integer())
                .c_str());
#else
        RequireAcceleratedBackend("ProfileFillStages");
#endif
    case DispatchId::BenchmarkPixelLoad:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(benchmarkPixelLoad(target.str(),
                                        arguments.At(1).integer(),
                                        arguments.At(2).integer(),
                                        arguments.At(3).integer())
                         .c_str());
#else
        RequireAcceleratedBackend("BenchmarkPixelLoad");
#endif
    case DispatchId::PixelTextureExperiment:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        auto pixels = arguments.At(1);
        if (pixels.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array of pixels");
        }
        return Value(pixelTextureExperiment(target.obj(),
                                            pixels.v.parray,
                                            arguments.At(2).integer(),
                                            arguments.At(3).integer(),
                                            arguments.At(4).integer(),
                                            arguments.At(5).integer())
                         .c_str());
    }
#else
        RequireAcceleratedBackend("PixelTextureExperiment");
#endif
    case DispatchId::ProbeShapePolicy:
        // Portable: the classification is ordinary Automation, and the fill
        // counter it reports exists on both backends.
        return Value(probeShapePolicy(target.obj()).c_str());
    case DispatchId::ProbeShapeCompatibility:
#if defined(BB_HAS_NATIVE_BACKEND)
        // Read-only classifier: never throws for an unsupported Shape class, so
        // the harness can put a row in the matrix instead of an exception.
        return Value(probeShapeCompatibility(target.obj()).c_str());
#else
        RequireAcceleratedBackend("ProbeShapeCompatibility");
#endif
    case DispatchId::ApplyPicture:
        return Value(applyPictureThroughAbi(target.obj(), arguments.At(1).str()).c_str());
    case DispatchId::InvalidateShape:
        return Value(invalidateShapeThroughAbi(target.obj()).c_str());
    case DispatchId::ApplyTextureUnrestricted:
#if defined(BB_HAS_NATIVE_BACKEND)
        return Value(applyTextureUnrestricted(target.obj(), arguments.At(1).integer()).c_str());
#else
        RequireAcceleratedBackend("ApplyTextureUnrestricted");
#endif
    case DispatchId::InspectFillReceiver:
#if defined(BB_HAS_NATIVE_BACKEND)
        // Throws bb::Error naming the failed guard; Invoke reports it verbatim.
        return Value(inspectFillReceiver(target.obj()).c_str());
#else
        RequireAcceleratedBackend("InspectFillReceiver");
#endif
    case DispatchId::LoadCachedImageExperiment:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        if (target.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(loadCachedImageExperiment(target.v.parray).c_str());
    }
#else
        RequireAcceleratedBackend("LoadCachedImageExperiment");
#endif
    case DispatchId::NativeApplyExperiment:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        auto imageBytes = arguments.At(1);
        if (imageBytes.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(nativeApplyExperiment(target.obj(), imageBytes.v.parray).c_str());
    }
#else
        RequireAcceleratedBackend("NativeApplyExperiment");
#endif
    case DispatchId::NativeApplyReuseExperiment:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        auto imageBytes = arguments.At(1);
        if (!(target.v.vt & VT_ARRAY) || !target.v.parray) {
            throw Error(E_INVALIDARG, "Expected an array of FillFormats");
        }
        if (imageBytes.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(nativeApplyReuseExperiment(target.v.parray, imageBytes.v.parray).c_str());
    }
#else
        RequireAcceleratedBackend("NativeApplyReuseExperiment");
#endif
    case DispatchId::TraceCachedApply:
#if defined(BB_HAS_NATIVE_BACKEND)
    {
        auto destination = arguments.At(1);
        check(traceCachedApply(target.obj(), destination.obj(), arguments.At(2).str()),
              "Cached trace");
        break;
    }
#else
        RequireAcceleratedBackend("TraceCachedApply");
#endif
    case DispatchId::LoadTextureBytesAbi: {
        if (target.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected a Byte array of encoded image data");
        }
        return Value(loadTextureBytesThroughAbi(target.v.parray).c_str());
    }
    case DispatchId::LoadTexturePixelsAbi: {
        if (target.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected a Byte array of BGRA pixels");
        }
        return Value(loadTexturePixelsThroughAbi(target.v.parray,
                                                 arguments.At(1).integer(),
                                                 arguments.At(2).integer(),
                                                 arguments.At(3).integer())
                         .c_str());
    }
    case DispatchId::LoadImagePixelsAbi: {
        if (target.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected a Byte array of BGRA pixels");
        }
        return Value(loadImagePixelsThroughAbi(target.v.parray,
                                               arguments.At(1).integer(),
                                               arguments.At(2).integer(),
                                               arguments.At(3).integer())
                         .c_str());
    }
    case DispatchId::ImageSizeAbi:
        return Value(imageSizeThroughAbi(target.integer()).c_str());
    case DispatchId::CreateTextureFromImageAbi:
        return Value(createTextureFromImageThroughAbi(target.integer()).c_str());
    case DispatchId::WarpImageQuadAbi: {
        auto points = arguments.At(1);
        if (points.v.vt != (VT_ARRAY | VT_R8) && points.v.vt != (VT_ARRAY | VT_VARIANT)) {
            throw Error(E_INVALIDARG, "Expected an array of eight quad coordinates");
        }
        return Value(warpImageQuadThroughAbi(
                         target.integer(), points.v.parray, arguments.At(2).integer())
                         .c_str());
    }
    case DispatchId::ApplyTextureBatchAbi: {
        if (!(target.v.vt & VT_ARRAY) || !target.v.parray) {
            throw Error(E_INVALIDARG, "Expected an array of Shapes");
        }
        return Value(
            applyTextureBatchThroughAbi(target.v.parray, arguments.At(1).integer()).c_str());
    }
    case DispatchId::AbiLifecycle:
        return Value(lifecycleThroughAbi(target.str()).c_str());
    default:
        throw Error(DISP_E_MEMBERNOTFOUND, "Unknown research operation");
    }
    return Value();
}
} // namespace bb
