/**
 * @file batch_benchmark.cpp
 * Measures BB_ApplyTextureBatch against the same number of BB_ApplyTexture calls.
 *
 * The batch exists to cut language-boundary crossings, so the honest question is
 * how much of the cost is the crossing and how much is the work. This runs both
 * legs through the real C ABI, in process, on the same Shapes with the same
 * texture in the same document.
 *
 * What this measures is the **ABI-side** difference: one entry, one thread check
 * and one argument validation instead of N. A VBA caller saves more than this,
 * because each `Declare` call also costs an interpreter-to-native transition
 * that is not present here - so treat the numbers below as a floor on the batch
 * benefit, not a ceiling. That extra saving is not measured and is not claimed.
 */

#include "../experiment_api.hpp"

#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>

#include <algorithm>
#include <sstream>
#include <vector>

namespace {

double SecondsPerTick() {
    LARGE_INTEGER frequency{};
    QueryPerformanceFrequency(&frequency);
    return 1.0 / static_cast<double>(frequency.QuadPart);
}

long long Now() {
    LARGE_INTEGER counter{};
    QueryPerformanceCounter(&counter);
    return counter.QuadPart;
}

struct Timing {
    double meanMs = 0.0;
    double medianMs = 0.0;
};

Timing Summarise(std::vector<double>& samples) {
    Timing timing;
    if (samples.empty()) {
        return timing;
    }
    double total = 0.0;
    for (double sample : samples) {
        total += sample;
    }
    timing.meanMs = total / samples.size();
    std::sort(samples.begin(), samples.end());
    timing.medianMs = samples[samples.size() / 2];
    return timing;
}

} // namespace

/**
 * Fills @p shapeCount Shapes @p iterations times each way and reports both.
 *
 * The Shapes are created and deleted here, so the caller's slide is left as it
 * was found.
 */
std::wstring benchmarkTextureBatch(IDispatch* slide, long shapeCount, long iterations) {
    if (shapeCount <= 0 || iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Shape count and iterations must be positive");
    }
    if (BB_Init() != BB_OK) {
        char message[512]{};
        BB_GetLastError(message, sizeof(message));
        throw bb::Error(E_NOTIMPL, std::string("BB_Init failed: ") + message);
    }

    bb::Value shapes = bb::get(slide, L"Shapes");
    std::vector<bb::Value> created;
    std::vector<void*> pointers;
    created.reserve(shapeCount);
    pointers.reserve(shapeCount);
    for (long index = 0; index < shapeCount; ++index) {
        // A grid that stays on the slide; geometry is irrelevant to the timing.
        const double left = 10.0 + static_cast<double>(index % 20) * 24.0;
        const double top = 10.0 + static_cast<double>(index / 20) * 24.0;
        bb::Value shape = bb::call(shapes.obj(), L"AddShape",
                                   {bb::Value(1L), bb::Value(left), bb::Value(top),
                                    bb::Value(20.0), bb::Value(20.0)});
        pointers.push_back(shape.obj());
        created.push_back(std::move(shape));
    }

    std::wostringstream out;
    BB_Handle texture = 0;
    try {
        // A 1x1 PNG keeps this self-contained. The image only has to decode;
        // decoding happens once, before either leg is timed.
        static const unsigned char kOnePixelPng[] = {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
            0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
            0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00,
            0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82};
        if (BB_LoadTexture(kOnePixelPng, sizeof(kOnePixelPng), &texture) != BB_OK) {
            char message[512]{};
            BB_GetLastError(message, sizeof(message));
            throw bb::Error(E_FAIL, std::string("BB_LoadTexture failed: ") + message);
        }

        std::vector<BB_Handle> handles(static_cast<std::size_t>(shapeCount), texture);
        const double tick = SecondsPerTick();
        std::vector<double> individual;
        std::vector<double> batched;
        individual.reserve(iterations);
        batched.reserve(iterations);

        // Warm both legs once so neither pays first-call setup in its samples.
        for (void* shape : pointers) {
            BB_ApplyTexture(shape, texture);
        }
        uint32_t applied = 0;
        BB_ApplyTextureBatch(pointers.data(), handles.data(),
                             static_cast<uint32_t>(shapeCount), &applied);

        for (long round = 0; round < iterations; ++round) {
            const long long start = Now();
            for (void* shape : pointers) {
                if (BB_ApplyTexture(shape, texture) != BB_OK) {
                    throw bb::Error(E_FAIL, "BB_ApplyTexture failed during the benchmark");
                }
            }
            individual.push_back((Now() - start) * tick * 1000.0);
        }

        for (long round = 0; round < iterations; ++round) {
            const long long start = Now();
            if (BB_ApplyTextureBatch(pointers.data(), handles.data(),
                                     static_cast<uint32_t>(shapeCount), &applied) != BB_OK) {
                throw bb::Error(E_FAIL, "BB_ApplyTextureBatch failed during the benchmark");
            }
            batched.push_back((Now() - start) * tick * 1000.0);
        }
        if (applied != static_cast<uint32_t>(shapeCount)) {
            throw bb::Error(E_FAIL, "The batch reported fewer applies than Shapes");
        }

        const Timing loop = Summarise(individual);
        const Timing batch = Summarise(batched);
        out.setf(std::ios::fixed);
        out.precision(4);
        out << L"shapes=" << shapeCount << L";iterations=" << iterations << L';'
            << L"individualMeanMs=" << loop.meanMs << L";individualMedianMs="
            << loop.medianMs << L';'
            << L"batchMeanMs=" << batch.meanMs << L";batchMedianMs=" << batch.medianMs
            << L';'
            << L"perShapeIndividualMs=" << (loop.meanMs / shapeCount)
            << L";perShapeBatchMs=" << (batch.meanMs / shapeCount) << L';';
        if (batch.meanMs > 0.0) {
            out << L"batchSpeedup=" << (loop.meanMs / batch.meanMs) << L';';
        }
        out << L"savedPerShapeUs=" << ((loop.meanMs - batch.meanMs) / shapeCount * 1000.0)
            << L';';
    } catch (...) {
        if (texture) {
            BB_ReleaseTexture(texture);
        }
        for (bb::Value& shape : created) {
            bb::call(shape.obj(), L"Delete");
        }
        throw;
    }

    BB_ReleaseTexture(texture);
    for (bb::Value& shape : created) {
        bb::call(shape.obj(), L"Delete");
    }
    return out.str();
}
