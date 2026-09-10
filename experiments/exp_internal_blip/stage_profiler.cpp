/**
 * @file stage_profiler.cpp
 * Stage-by-stage attribution of what a picture fill actually costs.
 *
 * Applying an already-created texture costs about 0.19 ms; creating new content
 * and applying it costs about 0.9-1.0 ms. Before trying to close that gap it has
 * to be known where the difference is, so this measures every step separately
 * rather than inferring it from two end-to-end totals.
 *
 * The measured steps are:
 *
 *  - `create`      GEL::ICachedImage::Create - decode or pixel upload
 *  - `resolve`     ResolveFillTarget plus ResolveApplyFunctions (the guards)
 *  - `record`      the OART property record constructor and its fill-kind slot
 *  - `subrecord`   the image sub-record constructor
 *  - `install`     putting the cached image into the sub-record (AddRef)
 *  - `transfer`    copying the sub-record into the record's image slot (AddRef)
 *  - `holder`      the stretch holder and the trailing tagged flag
 *  - `transaction` the OART transaction constructor
 *  - `apply`       the receiver call: document mutation, undo entry, invalidation
 *  - `released`    the four destructors, in the handler's own reverse order
 *  - `release`     releasing our own references to the created image
 *
 * The in-apply steps come from ApplyCachedImage's optional StageSampler, which
 * is not installed on any production call, so nothing here changes the behaviour
 * or the cost of the immutable path.
 *
 * ## What this cannot see
 *
 * `apply` includes whatever OART does synchronously - the document edit, the
 * undo record, marking the Shape dirty. It does **not** include repainting: this
 * loop never pumps messages, so deferred rendering happens after the
 * measurement. That is stated rather than papered over; a workload that yields
 * between frames pays it and these numbers do not show it.
 *
 * Four legs run, because the interesting quantity is a difference:
 *
 *  - `reuse`       one texture, applied repeatedly - the 0.19 ms case
 *  - `newSame`     create from the same encoded bytes every iteration
 *  - `newDistinct` create from raw pixels whose content changes every iteration
 *  - `pool`        a fixed ring of pre-created distinct textures, applied in turn
 *  - `abi`         the same new-content work through the real public C ABI
 *  - `overhead`    the individual guard and Automation primitives an apply pays
 *
 * The `abi` leg exists to settle where the previously reported 0.9-1.0 ms went.
 * That figure was taken from PowerShell, which pays a cross-process Automation
 * round trip per call and made three of them per frame. Running the identical
 * sequence in process through BB_LoadTexturePixels / BB_ApplyTexture /
 * BB_ReleaseTexture separates what a caller pays from what the harness paid.
 *
 * `newSame` against `newDistinct` matters because a cached image is
 * content-addressed: repeating identical bytes could be answered from Office's
 * own cache and would then flatter the create cost. `pool` is the bounded-content
 * ceiling - no creation on the hot path, but genuinely different content each
 * iteration - and is also where unbounded growth would show up.
 *
 * Research only. STA, PowerPoint host, and it leaves the document as found.
 */

#include "../../src/backend/windows_office/native_apply.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"
#include "../experiment_api.hpp"

#include <algorithm>
#include <blipbridge/blipbridge.h>
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstring>
#include <cwchar>
#include <filesystem>
#include <fstream>
#include <psapi.h>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

namespace {

constexpr long kDefaultIterations = 300;

/**
 * Ring size for the pool leg. Small on purpose: the question is whether a
 * *bounded* set of pre-created images can serve changing content, so a pool
 * large enough to hold every frame would be answering a different question.
 */
constexpr long kPoolSize = 8;

/// Raw-pixel frame size for the distinct-content legs. Large enough that
/// creation is not lost in noise, small enough to keep the run to seconds.
constexpr std::uint32_t kFrameWidth = 128;
constexpr std::uint32_t kFrameHeight = 128;
constexpr std::int32_t kFrameStride = static_cast<std::int32_t>(kFrameWidth) * 4;

/// Frame indices for the abi leg start here, so no frame it creates was already
/// created - and possibly cached - by an earlier leg.
constexpr long kAbiFrameOffset = 1000000;

/// The pool's frames get their own range for the same reason.
constexpr long kPoolFrameOffset = 2000000;

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

/// Nearest-rank percentile on an already sorted sample.
double Percentile(const std::vector<double>& sorted, double fraction) {
    if (sorted.empty()) {
        return 0.0;
    }
    const auto rank = static_cast<std::size_t>(fraction * static_cast<double>(sorted.size()));
    return sorted[std::min(rank, sorted.size() - 1)];
}

/**
 * One named stage's samples across one leg.
 *
 * The median is reported alongside the mean because a single scheduling stall in
 * a few-hundred-iteration run moves the mean of a ten-microsecond stage by more
 * than the stage itself costs.
 */
class Stage {
  public:
    void Add(double milliseconds) {
        samples_.push_back(milliseconds);
    }

    double MeanMs() const {
        if (samples_.empty()) {
            return 0.0;
        }
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        return total / static_cast<double>(samples_.size());
    }

    double MedianMs() const {
        return Quantile(0.50);
    }

    double P95Ms() const {
        return Quantile(0.95);
    }

  private:
    double Quantile(double fraction) const {
        std::vector<double> sorted = samples_;
        std::sort(sorted.begin(), sorted.end());
        return Percentile(sorted, fraction);
    }

    std::vector<double> samples_;
};

/**
 * The stages of one leg, held in first-seen order rather than sorted, so the
 * report reads in the same order as the code path it describes.
 */
class LegProfile {
  public:
    explicit LegProfile(const wchar_t* name) : name_(name) {}

    Stage& operator[](const wchar_t* stage) {
        for (auto& entry : stages_) {
            if (entry.first == stage) {
                return entry.second;
            }
        }
        stages_.emplace_back(stage, Stage{});
        return stages_.back().second;
    }

    /// Writes every stage plus the sum of the stage means, which is the leg's
    /// end-to-end cost broken into parts that add up.
    void Write(std::wostringstream& out) const {
        double totalMean = 0.0;
        for (const auto& [stage, samples] : stages_) {
            out << name_ << L'.' << stage << L".meanMs=" << samples.MeanMs() << L';' << name_
                << L'.' << stage << L".medianMs=" << samples.MedianMs() << L';' << name_ << L'.'
                << stage << L".p95Ms=" << samples.P95Ms() << L';';
            totalMean += samples.MeanMs();
        }
        out << name_ << L".totalMeanMs=" << totalMean << L';';
    }

  private:
    std::wstring name_;
    std::vector<std::pair<std::wstring, Stage>> stages_;
};

/**
 * Turns ApplyCachedImage's named checkpoints into per-stage durations.
 *
 * Each checkpoint's cost is the time since the previous one; `enter` establishes
 * the baseline and is not itself a stage.
 */
class ApplyTimer {
  public:
    ApplyTimer(LegProfile& profile, double tick) : profile_(profile), tick_(tick) {}

    bb::oart::StageSampler Sampler() {
        return [this](const wchar_t* stage) {
            const long long now = Now();
            if (std::wcscmp(stage, L"enter") == 0) {
                previous_ = now;
                return;
            }
            profile_[stage].Add(static_cast<double>(now - previous_) * tick_ * 1000.0);
            previous_ = now;
        };
    }

  private:
    LegProfile& profile_;
    double tick_ = 0.0;
    long long previous_ = 0;
};

struct ArrayGuard {
    SAFEARRAY* value;

    ~ArrayGuard() {
        SafeArrayDestroy(value);
    }
};

/// Owns the two references CreateCachedImageFrom* hands back, so an exception
/// anywhere in a leg still releases them.
class OwnedImage {
  public:
    OwnedImage() = default;

    explicit OwnedImage(bb::oart::CreatedImage created) : created_(created) {}

    OwnedImage(const OwnedImage&) = delete;
    OwnedImage& operator=(const OwnedImage&) = delete;

    OwnedImage(OwnedImage&& other) noexcept : created_(other.created_) {
        other.created_ = {};
    }

    OwnedImage& operator=(OwnedImage&& other) noexcept {
        if (this != &other) {
            Release();
            created_ = other.created_;
            other.created_ = {};
        }
        return *this;
    }

    ~OwnedImage() {
        Release();
    }

    void* cached() const {
        return created_.cached;
    }

    void Release() {
        // Same order as the texture store: the companion image first, then the
        // cached image Office holds while a fill is in place.
        bb::oart::ReleaseIntrusive(created_.image);
        bb::oart::ReleaseIntrusive(created_.cached);
        created_ = {};
    }

  private:
    bb::oart::CreatedImage created_{};
};

/**
 * A BGRA32 frame whose content genuinely differs per id, so a content-addressed
 * cache cannot answer two of them with the same image.
 *
 * The gradient alone would not be enough: every channel of it is computed modulo
 * 256, so ids 1,000,000 apart produce byte-identical pixels. The full 32-bit id
 * is therefore written into the top-left four pixels' colour channels, which
 * makes distinct ids distinct content by construction rather than by luck.
 */
std::vector<std::uint8_t> MakeFrame(long id) {
    std::vector<std::uint8_t> pixels(static_cast<std::size_t>(kFrameWidth) * kFrameHeight * 4);
    for (std::uint32_t y = 0; y < kFrameHeight; ++y) {
        for (std::uint32_t x = 0; x < kFrameWidth; ++x) {
            std::uint8_t* pixel =
                pixels.data() + (static_cast<std::size_t>(y) * kFrameWidth + x) * 4;
            pixel[0] = static_cast<std::uint8_t>(x + id);     // B
            pixel[1] = static_cast<std::uint8_t>(y + id * 3); // G
            pixel[2] = static_cast<std::uint8_t>(id);         // R
            pixel[3] = 0xFF;                                  // A
        }
    }
    const auto unique = static_cast<std::uint32_t>(id);
    for (std::uint32_t byte = 0; byte < 4; ++byte) {
        // Colour channels only; every alpha stays opaque so the frames remain
        // ordinary images rather than something Office might treat specially.
        std::uint8_t* pixel = pixels.data() + static_cast<std::size_t>(byte) * 4;
        pixel[byte % 3] = static_cast<std::uint8_t>((unique >> (byte * 8)) & 0xFF);
    }
    return pixels;
}

/// Private bytes, so the pool leg can be checked for unbounded growth.
std::size_t PrivateBytes() {
    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = sizeof(counters);
    if (!GetProcessMemoryInfo(GetCurrentProcess(),
                              reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters),
                              sizeof(counters))) {
        return 0;
    }
    return counters.PrivateUsage;
}

} // namespace

/**
 * Profiles the picture-fill path stage by stage on one Shape of @p slide.
 *
 * @p imagePath supplies the encoded bytes for the `newSame` leg; the distinct
 * legs synthesise their own pixels. The Shape is created and deleted here, so
 * the caller's document is left as it was found.
 */
std::wstring profileFillStages(IDispatch* slide, const std::wstring& imagePath, long iterations) {
    if (iterations <= 0) {
        iterations = kDefaultIterations;
    }
    if (!GetModuleHandleW(L"POWERPNT.EXE")) {
        throw bb::Error(E_ACCESSDENIED, "The stage profiler requires the PowerPoint host");
    }

    std::ifstream file(std::filesystem::path(imagePath), std::ios::binary);
    if (!file) {
        throw bb::Error(E_INVALIDARG, "Cannot read the profiler image");
    }
    const std::vector<char> contents((std::istreambuf_iterator<char>(file)),
                                     std::istreambuf_iterator<char>());
    if (contents.empty()) {
        throw bb::Error(E_INVALIDARG, "Profiler image is empty");
    }

    SAFEARRAYBOUND bound{static_cast<ULONG>(contents.size()), 0};
    SAFEARRAY* bytes = SafeArrayCreate(VT_UI1, 1, &bound);
    if (!bytes) {
        throw std::bad_alloc();
    }
    ArrayGuard bytesGuard{bytes};
    void* raw = nullptr;
    bb::check(SafeArrayAccessData(bytes, &raw), "SafeArrayAccessData");
    std::memcpy(raw, contents.data(), contents.size());
    SafeArrayUnaccessData(bytes);

    const double tick = SecondsPerTick();

    bb::Value shapes = bb::get(slide, L"Shapes");
    bb::Value shape = bb::call(
        shapes.obj(),
        L"AddShape",
        {bb::Value(1L), bb::Value(20.0), bb::Value(20.0), bb::Value(120.0), bb::Value(120.0)});

    LegProfile reuse(L"reuse");
    LegProfile newSame(L"newSame");
    LegProfile newDistinct(L"newDistinct");
    LegProfile pool(L"pool");
    LegProfile abi(L"abi");
    LegProfile overhead(L"overhead");
    std::size_t privateBefore = 0;
    std::size_t privateAfter = 0;
    std::wstring poolCounts;

    try {
        bb::Value fill = bb::get(shape.obj(), L"Fill");
        IDispatch* const fillObject = fill.obj();

        // Warm every path once: the first call of each pays one-off setup that
        // would otherwise land entirely in the first sample of a stage.
        {
            OwnedImage warm(bb::oart::CreateCachedImageFromBytes(bytes));
            const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
            const bb::oart::ApplyFunctions functions =
                bb::oart::ResolveApplyFunctions(target.oartBase);
            bb::oart::ApplyCachedImage(functions, target, warm.cached());
        }

        // --- reuse: one image, applied repeatedly. The 0.19 ms case.
        {
            OwnedImage texture(bb::oart::CreateCachedImageFromBytes(bytes));
            ApplyTimer timer(reuse, tick);
            const auto sampler = timer.Sampler();
            for (long index = 0; index < iterations; ++index) {
                const long long resolveStart = Now();
                const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
                const bb::oart::ApplyFunctions functions =
                    bb::oart::ResolveApplyFunctions(target.oartBase);
                reuse[L"resolve"].Add(static_cast<double>(Now() - resolveStart) * tick * 1000.0);
                bb::oart::ApplyCachedImage(functions, target, texture.cached(), sampler);
            }
        }

        // --- newSame: identical bytes re-created every iteration. If Office
        // answers from its own content cache, `create` here lands far below the
        // newDistinct row - which is a finding, not a flaw in the test.
        {
            ApplyTimer timer(newSame, tick);
            const auto sampler = timer.Sampler();
            for (long index = 0; index < iterations; ++index) {
                const long long createStart = Now();
                OwnedImage texture(bb::oart::CreateCachedImageFromBytes(bytes));
                newSame[L"create"].Add(static_cast<double>(Now() - createStart) * tick * 1000.0);

                const long long resolveStart = Now();
                const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
                const bb::oart::ApplyFunctions functions =
                    bb::oart::ResolveApplyFunctions(target.oartBase);
                newSame[L"resolve"].Add(static_cast<double>(Now() - resolveStart) * tick * 1000.0);

                bb::oart::ApplyCachedImage(functions, target, texture.cached(), sampler);

                const long long releaseStart = Now();
                texture.Release();
                newSame[L"release"].Add(static_cast<double>(Now() - releaseStart) * tick * 1000.0);
            }
        }

        // --- newDistinct: genuinely new content every iteration, which is what
        // a changing picture actually generates. Frame synthesis stays outside
        // the timed region: it is the caller's cost, not Office's.
        {
            ApplyTimer timer(newDistinct, tick);
            const auto sampler = timer.Sampler();
            for (long index = 0; index < iterations; ++index) {
                const std::vector<std::uint8_t> frame = MakeFrame(index);

                const long long createStart = Now();
                OwnedImage texture(bb::oart::CreateCachedImageFromPixels(
                    frame.data(), kFrameWidth, kFrameHeight, kFrameStride));
                newDistinct[L"create"].Add(static_cast<double>(Now() - createStart) * tick *
                                           1000.0);

                const long long resolveStart = Now();
                const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
                const bb::oart::ApplyFunctions functions =
                    bb::oart::ResolveApplyFunctions(target.oartBase);
                newDistinct[L"resolve"].Add(static_cast<double>(Now() - resolveStart) * tick *
                                            1000.0);

                bb::oart::ApplyCachedImage(functions, target, texture.cached(), sampler);

                const long long releaseStart = Now();
                texture.Release();
                newDistinct[L"release"].Add(static_cast<double>(Now() - releaseStart) * tick *
                                            1000.0);
            }
        }

        // --- pool: a bounded ring of pre-created distinct images, applied in
        // turn with no creation on the hot path. This is the ceiling available
        // to any pooling scheme, and the place unbounded growth would show.
        {
            std::vector<OwnedImage> ring;
            ring.reserve(kPoolSize);
            for (long index = 0; index < kPoolSize; ++index) {
                const std::vector<std::uint8_t> frame = MakeFrame(kPoolFrameOffset + index);
                ring.emplace_back(bb::oart::CreateCachedImageFromPixels(
                    frame.data(), kFrameWidth, kFrameHeight, kFrameStride));
            }
            privateBefore = PrivateBytes();
            ApplyTimer timer(pool, tick);
            const auto sampler = timer.Sampler();
            for (long index = 0; index < iterations; ++index) {
                const long long resolveStart = Now();
                const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fillObject);
                const bb::oart::ApplyFunctions functions =
                    bb::oart::ResolveApplyFunctions(target.oartBase);
                pool[L"resolve"].Add(static_cast<double>(Now() - resolveStart) * tick * 1000.0);
                bb::oart::ApplyCachedImage(
                    functions,
                    target,
                    ring[static_cast<std::size_t>(index % kPoolSize)].cached(),
                    sampler);
            }
            privateAfter = PrivateBytes();

            // Every ring entry's reference count after the run. An image the
            // document retained per apply would show a count that climbed with
            // the iteration count instead of sitting at its resting value.
            std::wostringstream counts;
            for (const OwnedImage& entry : ring) {
                counts << bb::oart::IntrusiveCount(entry.cached()) << L',';
            }
            poolCounts = counts.str();
        }

        // --- abi: the same new-content frame through the public entry points,
        // in process. Anything above the newDistinct total is the handle store
        // and the ABI itself; anything a PowerShell run shows above *this* is
        // the Automation round trip, not work.
        if (BB_Init() == BB_OK) {
            IDispatch* const shapeObject = shape.obj();
            for (long index = 0; index < iterations; ++index) {
                // Offset past every frame the newDistinct leg already created.
                // Repeating that content would let Office's content-addressed
                // cache answer the create, and this leg would then be timing a
                // cache hit while claiming to time a creation.
                const std::vector<std::uint8_t> frame = MakeFrame(index + kAbiFrameOffset);

                BB_Handle handle = 0;
                const long long loadStart = Now();
                const BB_Result loaded = BB_LoadTexturePixels(
                    frame.data(), kFrameWidth, kFrameHeight, kFrameStride, &handle);
                abi[L"BB_LoadTexturePixels"].Add(static_cast<double>(Now() - loadStart) * tick *
                                                 1000.0);
                if (loaded != BB_OK) {
                    throw bb::Error(E_FAIL, "BB_LoadTexturePixels failed during the profile");
                }

                const long long applyStart = Now();
                const BB_Result applied = BB_ApplyTexture(shapeObject, handle);
                abi[L"BB_ApplyTexture"].Add(static_cast<double>(Now() - applyStart) * tick *
                                            1000.0);

                const long long releaseStart = Now();
                BB_ReleaseTexture(handle);
                abi[L"BB_ReleaseTexture"].Add(static_cast<double>(Now() - releaseStart) * tick *
                                              1000.0);
                if (applied != BB_OK) {
                    throw bb::Error(E_FAIL, "BB_ApplyTexture failed during the profile");
                }
            }
            // Deliberately no BB_Shutdown: the C ABI and the COM Engine share
            // one texture store, so shutting the library down here would also
            // release handles the Engine had handed out. Each handle this leg
            // creates is released in the loop, which is all that is owed.
        }

        // --- overhead: the individual pieces every apply pays before any Office
        // work happens. `resolve` and the ABI's Shape validation are both ours,
        // so unlike the receiver call they are things this project could change;
        // measuring them separately is what says whether that would be worth it.
        //
        // These are the same primitives the backend calls, in the same order,
        // rather than the backend itself: instrumenting the production path would
        // mean putting timers in it.
        {
            IDispatch* const shapeObject = shape.obj();
            for (long index = 0; index < iterations; ++index) {
                const long long typeStart = Now();
                const long shapeType = bb::get(shapeObject, L"Type").integer();
                overhead[L"getType"].Add(static_cast<double>(Now() - typeStart) * tick * 1000.0);
                if (shapeType == 0) {
                    throw bb::Error(E_FAIL, "Shape reported no type during the profile");
                }

                const long long fillStart = Now();
                bb::Value fetched = bb::get(shapeObject, L"Fill");
                overhead[L"getFill"].Add(static_cast<double>(Now() - fillStart) * tick * 1000.0);

                const long long targetStart = Now();
                const bb::oart::FillTarget target = bb::oart::ResolveFillTarget(fetched.obj());
                overhead[L"resolveTarget"].Add(static_cast<double>(Now() - targetStart) * tick *
                                               1000.0);

                const long long functionsStart = Now();
                const bb::oart::ApplyFunctions functions =
                    bb::oart::ResolveApplyFunctions(target.oartBase);
                overhead[L"resolveFunctions"].Add(static_cast<double>(Now() - functionsStart) *
                                                  tick * 1000.0);
                if (!functions.constructRecord) {
                    throw bb::Error(E_FAIL, "Apply functions did not resolve during the profile");
                }

                // ResolveFillTarget turned out to be nearly all of the guard
                // cost, so its own ingredients are timed individually. Whichever
                // of these dominates is the only thing worth reconsidering, and
                // reconsidering a guard means proving it stays as strict.
                const long long readableStart = Now();
                const bool readable = bb::oart::IsReadable(fetched.obj(), 0x10);
                overhead[L"isReadable"].Add(static_cast<double>(Now() - readableStart) * tick *
                                            1000.0);
                if (!readable) {
                    throw bb::Error(E_FAIL, "FillFormat became unreadable during the profile");
                }

                const long long handleStart = Now();
                const HMODULE oart = GetModuleHandleW(L"oart.dll");
                overhead[L"getModuleHandle"].Add(static_cast<double>(Now() - handleStart) * tick *
                                                 1000.0);

                const long long infoStart = Now();
                MODULEINFO moduleInfo{};
                GetModuleInformation(GetCurrentProcess(),
                                     GetModuleHandleW(L"ppcore.dll"),
                                     &moduleInfo,
                                     sizeof(moduleInfo));
                overhead[L"getModuleInformation"].Add(static_cast<double>(Now() - infoStart) *
                                                      tick * 1000.0);

                const long long supportedStart = Now();
                bb::oart::RequireSupportedModule(L"oart.dll", "OART");
                overhead[L"requireSupportedModule"].Add(
                    static_cast<double>(Now() - supportedStart) * tick * 1000.0);

                const long long wrapperStart = Now();
                bb::oart::DelegatingWrapper wrapper;
                bb::oart::DescribeDelegatingWrapper(
                    fetched.obj(),
                    reinterpret_cast<std::uintptr_t>(GetModuleHandleW(L"ppcore.dll")),
                    moduleInfo.SizeOfImage,
                    wrapper);
                overhead[L"describeWrapper"].Add(static_cast<double>(Now() - wrapperStart) * tick *
                                                 1000.0);
                if (!oart) {
                    throw bb::Error(E_FAIL, "OART vanished during the profile");
                }
            }
        }
    } catch (...) {
        bb::call(shape.obj(), L"Delete");
        throw;
    }

    bb::call(shape.obj(), L"Delete");

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(4);
    out << L"iterations=" << iterations << L";frame=" << kFrameWidth << L'x' << kFrameHeight
        << L";poolSize=" << kPoolSize << L';';
    reuse.Write(out);
    newSame.Write(out);
    newDistinct.Write(out);
    pool.Write(out);
    abi.Write(out);
    overhead.Write(out);
    out << L"pool.cachedCounts=" << poolCounts << L';' << L"pool.privateBytesBefore="
        << privateBefore << L';' << L"pool.privateBytesAfter=" << privateAfter << L';'
        << L"pool.privateBytesDelta="
        << static_cast<long long>(privateAfter) - static_cast<long long>(privateBefore) << L';';
    return out.str();
}
