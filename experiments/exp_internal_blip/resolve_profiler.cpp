/**
 * @file resolve_profiler.cpp
 * Where the receiver resolution's ~23 microseconds go.
 *
 * `ResolveFillTarget` is the second largest cost in an apply after Office's own
 * document edit - 10% of the total - and unlike that edit it is entirely ours,
 * so it is worth knowing which of its steps is expensive before trying to make
 * any of them cheaper.
 *
 * The walk is repeated here step by step rather than instrumented in place: the
 * production resolve must stay free of hooks, and a sampler threaded through it
 * would cost more than some of the steps being measured. That means this file
 * has to be kept in step with oart_layout.cpp by hand, which is acceptable for a
 * measurement that exists to answer one question and is re-run when it changes.
 *
 * The last leg is a bare VirtualQuery on an address already known to be
 * readable, because IsReadable is built out of VirtualQuery and the interesting
 * question is how much of the walk is that one call - it costs about 0.7 us in
 * an ordinary process, and PowerPoint's address space is not ordinary.
 */

#include "../experiment_api.hpp"

#include "../../src/backend/windows_office/oart_layout.hpp"

#include <algorithm>
#include <blipbridge/dispatch.hpp>
#include <sstream>
#include <string>
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

/// One step's samples, in microseconds. Reserved once; never formatted in loop.
class Step {
  public:
    Step(const wchar_t* name, long capacity) : name_(name) {
        samples_.reserve(static_cast<std::size_t>(capacity));
    }

    void Add(double us) {
        samples_.push_back(us);
    }

    double MeanUs() const {
        if (samples_.empty()) {
            return 0.0;
        }
        double total = 0.0;
        for (double sample : samples_) {
            total += sample;
        }
        return total / static_cast<double>(samples_.size());
    }

    void Write(std::wostringstream& out) {
        if (samples_.empty()) {
            return;
        }
        std::sort(samples_.begin(), samples_.end());
        out << L"resolve." << name_ << L".mean=" << MeanUs() << L';' << L"resolve." << name_
            << L".median=" << samples_[samples_.size() / 2] << L';' << L"resolve." << name_
            << L".min=" << samples_.front() << L';';
    }

  private:
    const wchar_t* name_;
    std::vector<double> samples_;
};

} // namespace

/**
 * Times each step of the receiver resolution for @p fill.
 *
 * @p fill is a Shape's FillFormat, exactly as the production path receives it.
 * Read-only: nothing is applied and no document state changes.
 */
std::wstring profileResolveStages(IDispatch* fill, long iterations) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing FillFormat");
    }
    if (iterations <= 0) {
        throw bb::Error(E_INVALIDARG, "Iterations must be positive");
    }

    using namespace bb::oart;

    // Layout constants, mirrored from oart_layout.cpp. Wrong values here would
    // show up as a thrown resolve, not as a quietly wrong measurement.
    constexpr std::size_t kPublicFillFormatSize = 0x10;
    constexpr std::size_t kFillFormatSize = 0x68;
    constexpr std::size_t kControlBlockSize = 0x18;
    constexpr std::size_t kReceiverInspectedSize = 0x28;
    constexpr std::size_t kFillFormatTokenOffset = 0x58;
    constexpr std::size_t kControlBlockPointeeOffset = 0x10;

    // Warm every cache the production path warms, so the samples measure the
    // steady state rather than first-call analysis.
    const FillTarget warm = ResolveFillTarget(fill);

    const double tick = SecondsPerTick();
    Step modules(L"synchroniseModules", iterations);
    Step validate(L"validateModules", iterations);
    Step imageSize(L"imageSize", iterations);
    Step readableFill(L"readableFill", iterations);
    Step wrapper(L"describeWrapper", iterations);
    Step readableHandler(L"readableHandler", iterations);
    Step readableToken(L"readableToken", iterations);
    Step readableReceiver(L"readableReceiver", iterations);
    Step bareQuery(L"bareVirtualQuery", iterations);
    Step whole(L"wholeResolve", iterations);
    std::vector<std::uintptr_t> regions;
    std::vector<std::size_t> regionSizes;

    for (long round = 0; round < iterations; ++round) {
        long long start = Now();
        const OfficeModules found = SynchroniseOfficeModules();
        modules.Add(static_cast<double>(Now() - start) * tick * 1e6);

        start = Now();
        const auto oartBase =
            reinterpret_cast<std::uintptr_t>(RequireValidatedModule(found.oart, "OART"));
        const auto ppcoreBase =
            reinterpret_cast<std::uintptr_t>(RequireValidatedModule(found.ppcore, "PPCORE"));
        validate.Add(static_cast<double>(Now() - start) * tick * 1e6);
        (void)oartBase;

        start = Now();
        const std::size_t ppcoreSize = OfficeModuleImageSize(found.ppcore);
        imageSize.Add(static_cast<double>(Now() - start) * tick * 1e6);

        start = Now();
        const bool fillReadable = IsReadable(fill, kPublicFillFormatSize);
        readableFill.Add(static_cast<double>(Now() - start) * tick * 1e6);
        if (!fillReadable) {
            throw bb::Error(E_FAIL, "FillFormat stopped being readable mid-profile");
        }

        DelegatingWrapper described{};
        start = Now();
        const bool isWrapper = DescribeDelegatingWrapper(fill, ppcoreBase, ppcoreSize, described);
        wrapper.Add(static_cast<double>(Now() - start) * tick * 1e6);
        if (!isWrapper) {
            throw bb::Error(E_FAIL, "The FillFormat stopped looking like a wrapper mid-profile");
        }

        auto handler = reinterpret_cast<void*>(LoadPointer(fill, described.innerOffset));
        start = Now();
        const bool handlerReadable = IsReadable(handler, kFillFormatSize);
        readableHandler.Add(static_cast<double>(Now() - start) * tick * 1e6);

        auto token = reinterpret_cast<void*>(LoadPointer(handler, kFillFormatTokenOffset));
        start = Now();
        const bool tokenReadable = IsReadable(token, kControlBlockSize);
        readableToken.Add(static_cast<double>(Now() - start) * tick * 1e6);

        auto receiver = reinterpret_cast<void*>(LoadPointer(token, kControlBlockPointeeOffset));
        start = Now();
        const bool receiverReadable = IsReadable(receiver, kReceiverInspectedSize);
        readableReceiver.Add(static_cast<double>(Now() - start) * tick * 1e6);
        if (!handlerReadable || !tokenReadable || !receiverReadable) {
            throw bb::Error(E_FAIL, "The receiver chain stopped being readable mid-profile");
        }

        MEMORY_BASIC_INFORMATION information{};
        start = Now();
        VirtualQuery(receiver, &information, sizeof(information));
        bareQuery.Add(static_cast<double>(Now() - start) * tick * 1e6);

        // Whether the four objects share memory regions decides whether one
        // confirmation could stand for several. Recorded once: it is a property
        // of the chain, not of the round.
        if (round == 0) {
            const void* chain[] = {fill, handler, token, receiver};
            for (const void* member : chain) {
                MEMORY_BASIC_INFORMATION region{};
                VirtualQuery(member, &region, sizeof(region));
                regions.push_back(reinterpret_cast<std::uintptr_t>(region.BaseAddress));
                regionSizes.push_back(region.RegionSize);
            }
        }

        // The production call itself, for comparison: the parts above should add
        // up to it, and a gap would mean this file has drifted from it.
        start = Now();
        const FillTarget resolved = ResolveFillTarget(fill);
        whole.Add(static_cast<double>(Now() - start) * tick * 1e6);
        if (resolved.receiver != warm.receiver) {
            throw bb::Error(E_FAIL, "The receiver moved between resolves");
        }
    }

    std::wostringstream out;
    out.setf(std::ios::fixed);
    out.precision(4);
    out << L"iterations=" << iterations << L';';
    modules.Write(out);
    validate.Write(out);
    imageSize.Write(out);
    readableFill.Write(out);
    wrapper.Write(out);
    readableHandler.Write(out);
    readableToken.Write(out);
    readableReceiver.Write(out);
    bareQuery.Write(out);
    whole.Write(out);

    const double parts = modules.MeanUs() + validate.MeanUs() + imageSize.MeanUs() +
                         readableFill.MeanUs() + wrapper.MeanUs() + readableHandler.MeanUs() +
                         readableToken.MeanUs() + readableReceiver.MeanUs();
    out << L"resolve.partsSum=" << parts << L';' << L"resolve.unattributed="
        << (whole.MeanUs() - parts) << L';';

    // How many distinct regions the four chain objects occupy: one confirmation
    // per region is the floor a per-call memo could reach.
    std::vector<std::uintptr_t> distinct = regions;
    std::sort(distinct.begin(), distinct.end());
    distinct.erase(std::unique(distinct.begin(), distinct.end()), distinct.end());
    out << L"resolve.chainRegions=" << distinct.size() << L';';
    for (std::size_t index = 0; index < regions.size(); ++index) {
        out << L"resolve.region" << index << L"=0x" << std::hex << regions[index] << L"+0x"
            << regionSizes[index] << std::dec << L';';
    }
    return out.str();
}
