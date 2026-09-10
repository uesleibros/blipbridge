/**
 * @file apply_cost_factors.cpp
 * What makes the private apply cost what it costs.
 *
 * The stage profiler attributed 76% of a cached apply to a single call - the
 * receiver's own handler, which performs the document edit. That number says
 * where the time goes but not *what it is*, and the difference matters: if the
 * bulk is rendering the changed Shape, then it is avoidable by not making the
 * Shape visible while it is being changed, and a batch is worth building. If the
 * bulk is document bookkeeping - undo, dirty flags, property storage - then it
 * is proportional to the number of edits and no arrangement of the slide will
 * help.
 *
 * So this varies exactly one thing at a time around the same apply and reports
 * what each is worth:
 *
 *   - the Shape is on the slide being displayed (the baseline)
 *   - the Shape is on a different slide, never shown
 *   - the Shape is hidden
 *   - the Shape is positioned off the slide area
 *   - the application window is minimised
 *
 * Every leg applies the same image to a Shape of the same size through the same
 * C ABI, and nothing about the fill differs. The document is left as found.
 */

#include "../experiment_api.hpp"

#include <algorithm>
#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>
#include <sstream>
#include <string>
#include <vector>

namespace {

constexpr long kMsoTrue = -1;
constexpr long kMsoFalse = 0;
constexpr long kPpLayoutBlank = 12;
/// PpWindowState: minimised, maximised.
constexpr long kPpWindowMinimized = 2;
constexpr long kPpWindowMaximized = 3;

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

[[noreturn]] void Fail(const char* what) {
    char message[512]{};
    BB_GetLastError(message, sizeof(message));
    throw bb::Error(E_FAIL, std::string(what) + ": " + message);
}

/// A 1x1 PNG: the image content is irrelevant here, only that it decodes once.
const unsigned char kPixelPng[] = {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
    0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
    0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
    0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D,
    0xB0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82};

/// Mean and median of one leg, sorted in place.
struct Result {
    double meanMs = 0.0;
    double medianMs = 0.0;
    double minMs = 0.0;
};

Result Measure(void* shape, BB_Handle texture, long iterations, double tick) {
    std::vector<double> samples;
    samples.reserve(static_cast<std::size_t>(iterations));
    // One untimed apply so the leg does not pay for whatever the previous leg's
    // rearrangement invalidated.
    if (BB_ApplyTexture(shape, texture) != BB_OK) {
        Fail("BB_ApplyTexture failed while warming a leg");
    }
    for (long round = 0; round < iterations; ++round) {
        const long long start = Now();
        const BB_Result applied = BB_ApplyTexture(shape, texture);
        const double ms = static_cast<double>(Now() - start) * tick * 1000.0;
        if (applied != BB_OK) {
            Fail("BB_ApplyTexture failed during a leg");
        }
        samples.push_back(ms);
    }
    std::sort(samples.begin(), samples.end());
    double total = 0.0;
    for (double sample : samples) {
        total += sample;
    }
    Result result;
    result.meanMs = total / static_cast<double>(samples.size());
    result.medianMs = samples[samples.size() / 2];
    result.minMs = samples.front();
    return result;
}

void Write(std::wostringstream& out, const wchar_t* name, const Result& result) {
    out << name << L"MeanMs=" << result.meanMs << L';' << name << L"MedianMs=" << result.medianMs
        << L';' << name << L"MinMs=" << result.minMs << L';';
}

} // namespace

/**
 * Measures the apply under five arrangements of the same Shape.
 *
 * @p presentation is used to add a second slide, which is removed again. The
 * window state is restored before returning, including after a failure.
 */
std::wstring measureApplyCostFactors(IDispatch* presentation, long iterations) {
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }
    if (BB_Init() != BB_OK) {
        Fail("BB_Init failed");
    }
    const double tick = SecondsPerTick();

    bb::Value slides = bb::get(presentation, L"Slides");
    bb::Value windows = bb::get(presentation, L"Windows");
    bb::Value window = bb::call(windows.obj(), L"Item", {bb::Value(1L)});
    const long originalState = bb::get(window.obj(), L"WindowState").integer();

    bb::Value firstSlide = bb::call(slides.obj(), L"Item", {bb::Value(1L)});
    bb::Value firstShapes = bb::get(firstSlide.obj(), L"Shapes");
    bb::Value visible = bb::call(firstShapes.obj(),
                                 L"AddShape",
                                 {bb::Value(1L),
                                  bb::Value(40.0),
                                  bb::Value(40.0),
                                  bb::Value(200.0),
                                  bb::Value(200.0)});

    // A second slide, added at the end so the displayed slide does not change.
    const long slideCount = bb::get(slides.obj(), L"Count").integer();
    bb::Value otherSlide =
        bb::call(slides.obj(), L"Add", {bb::Value(slideCount + 1), bb::Value(kPpLayoutBlank)});
    bb::Value otherShapes = bb::get(otherSlide.obj(), L"Shapes");
    bb::Value offSlideView = bb::call(otherShapes.obj(),
                                      L"AddShape",
                                      {bb::Value(1L),
                                       bb::Value(40.0),
                                       bb::Value(40.0),
                                       bb::Value(200.0),
                                       bb::Value(200.0)});

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(5);
    BB_Handle texture = 0;
    try {
        texture = 0;
        if (BB_LoadTexture(kPixelPng, sizeof(kPixelPng), &texture) != BB_OK) {
            Fail("BB_LoadTexture failed");
        }

        // Make sure the window really is showing slide 1, so "displayed" means
        // what it says.
        bb::Value view = bb::get(window.obj(), L"View");
        bb::call(view.obj(), L"GotoSlide", {bb::Value(1L)});
        bb::put(window.obj(), L"WindowState", bb::Value(kPpWindowMaximized));

        out << L"iterations=" << iterations << L';';

        Write(out, L"displayed", Measure(visible.obj(), texture, iterations, tick));
        Write(out, L"otherSlide", Measure(offSlideView.obj(), texture, iterations, tick));

        bb::put(visible.obj(), L"Visible", bb::Value(kMsoFalse));
        Write(out, L"hidden", Measure(visible.obj(), texture, iterations, tick));
        bb::put(visible.obj(), L"Visible", bb::Value(kMsoTrue));

        // Off the slide area entirely: still on the displayed slide, still
        // visible, but nothing of it is on screen.
        const double originalLeft = bb::get(visible.obj(), L"Left").number();
        bb::put(visible.obj(), L"Left", bb::Value(-4000.0));
        Write(out, L"offSlide", Measure(visible.obj(), texture, iterations, tick));
        bb::put(visible.obj(), L"Left", bb::Value(originalLeft));

        bb::put(window.obj(), L"WindowState", bb::Value(kPpWindowMinimized));
        Write(out, L"minimised", Measure(visible.obj(), texture, iterations, tick));
        bb::put(window.obj(), L"WindowState", bb::Value(originalState));
    } catch (...) {
        try {
            bb::put(window.obj(), L"WindowState", bb::Value(originalState));
            bb::put(visible.obj(), L"Visible", bb::Value(kMsoTrue));
        } catch (...) {
        }
        if (texture) {
            BB_ReleaseTexture(texture);
        }
        try {
            bb::call(visible.obj(), L"Delete");
            bb::call(otherSlide.obj(), L"Delete");
        } catch (...) {
        }
        throw;
    }

    BB_ReleaseTexture(texture);
    bb::call(visible.obj(), L"Delete");
    bb::call(otherSlide.obj(), L"Delete");
    return out.str();
}
