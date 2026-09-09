#include "experiment_api.hpp"
#include "../src/com/engine.hpp"

namespace bb {
namespace {
/// Argument counts are part of the research ABI; keep them in one place.
UINT ExpectedResearchArgumentCount(DispatchId id) {
    switch (id) {
    case DispatchId::RunBenchmarks:
    case DispatchId::InspectFillReceiver:
    case DispatchId::LoadCachedImageExperiment:
    case DispatchId::InspectTexture:
        return 1;
    case DispatchId::PixelTextureExperiment:
        return 6;
    case DispatchId::BenchmarkPixelLoad:
        return 4;
    case DispatchId::MemoryFillExperiment:
    case DispatchId::TraceCachedApply:
    case DispatchId::BenchmarkNativeTexture:
    case DispatchId::BenchmarkTextureBatch:
        return 3;
    default:
        return 2;
    }
}
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
        return Value(nativeTextureReport(target.integer()).c_str());
    case DispatchId::BenchmarkNativeTexture:
        return Value(benchmarkNativeTexture(target.obj(), arguments.At(1).str(),
                                            arguments.At(2).integer())
                         .c_str());
    case DispatchId::BenchmarkTextureBatch:
        return Value(benchmarkTextureBatch(target.obj(), arguments.At(1).integer(),
                                           arguments.At(2).integer())
                         .c_str());
    case DispatchId::BenchmarkPixelLoad:
        return Value(benchmarkPixelLoad(target.str(), arguments.At(1).integer(),
                                        arguments.At(2).integer(),
                                        arguments.At(3).integer())
                         .c_str());
    case DispatchId::PixelTextureExperiment: {
        auto pixels = arguments.At(1);
        if (pixels.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array of pixels");
        }
        return Value(pixelTextureExperiment(target.obj(), pixels.v.parray,
                                            arguments.At(2).integer(),
                                            arguments.At(3).integer(),
                                            arguments.At(4).integer(),
                                            arguments.At(5).integer())
                         .c_str());
    }
    case DispatchId::InspectFillReceiver:
        // Throws bb::Error naming the failed guard; Invoke reports it verbatim.
        return Value(inspectFillReceiver(target.obj()).c_str());
    case DispatchId::LoadCachedImageExperiment: {
        if (target.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(loadCachedImageExperiment(target.v.parray).c_str());
    }
    case DispatchId::NativeApplyExperiment: {
        auto imageBytes = arguments.At(1);
        if (imageBytes.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(nativeApplyExperiment(target.obj(), imageBytes.v.parray).c_str());
    }
    case DispatchId::NativeApplyReuseExperiment: {
        auto imageBytes = arguments.At(1);
        if (!(target.v.vt & VT_ARRAY) || !target.v.parray) {
            throw Error(E_INVALIDARG, "Expected an array of FillFormats");
        }
        if (imageBytes.v.vt != (VT_ARRAY | VT_UI1)) {
            throw Error(E_INVALIDARG, "Expected Byte array");
        }
        return Value(
            nativeApplyReuseExperiment(target.v.parray, imageBytes.v.parray).c_str());
    }
    case DispatchId::TraceCachedApply: {
        auto destination = arguments.At(1);
        check(traceCachedApply(target.obj(), destination.obj(), arguments.At(2).str()),
              "Cached trace");
        break;
    }
    default:
        throw Error(DISP_E_MEMBERNOTFOUND, "Unknown research operation");
    }
    return Value();
}
} // namespace bb
