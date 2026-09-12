/**
 * @file resolver.cpp
 * The decision. See the header for the order and for why every branch that is
 * not a clear yes ends at the portable backend.
 */

#include "resolver.hpp"

#include "validation_cache.hpp"
#include <sstream>

namespace bb::native {
namespace {

Resolution Portable(std::wstring reason) {
    Resolution resolution;
    resolution.decision = Decision::Portable;
    resolution.reason = std::move(reason);
    return resolution;
}

} // namespace

Resolution Resolve(const ArchitectureResolver& resolver,
                   const std::function<bool(const NativeProfile&, std::wstring&)>& selfTest) {
    /*
     * Step 1: the modules have to be there. GFX is delay-loaded in PowerPoint and
     * is legitimately absent until the first picture operation, so a missing
     * module here is an ordinary "not yet" rather than a fault - and it still
     * means no native backend for now.
     */
    ModuleWorld world;
    std::wstring missing;
    for (const wchar_t* name : resolver.RequiredModules()) {
        if (!world.Add(name)) {
            if (!missing.empty()) {
                missing += L", ";
            }
            missing += name;
        }
    }
    if (!missing.empty()) {
        return Portable(L"required Office modules are not loaded: " + missing);
    }

    /*
     * Step 2: an exact compiled profile, if one matches. Taken in preference to
     * anything resolved, and without consulting the resolver at all - a profile
     * derived offline against this exact build is the strongest evidence
     * available and there is nothing to gain by re-deriving it.
     */
    std::wostringstream rejected;
    for (const NativeProfile& candidate : resolver.ExactProfiles()) {
        std::wstring mismatch;
        if (!ProfileMatchesLoadedModules(candidate, mismatch)) {
            rejected << L"\n  - " << mismatch;
            continue;
        }
        if (!candidate.facts->Revalidate()) {
            /*
             * The identities matched but the memory did not, and that is a
             * different kind of failure from an unknown build. Every module this
             * profile names reports exactly the build it was derived against,
             * and the memory still does not look the way it should - so
             * something is loaded that is not what it says it is: a patched
             * image, a hook, or corruption.
             *
             * It refuses outright rather than continuing to structural
             * resolution. Resolving would mean inferring the layout of a binary
             * already caught misrepresenting itself, and a passing self-test
             * would not make that safe.
             */
            std::wostringstream out;
            out << L"exact profile matched the loaded modules but did not revalidate: "
                << candidate.Describe()
                << L". Something is loaded that is not the build it reports.";
            return Portable(out.str());
        }

        /*
         * An exact profile still has to prove itself once, and the reason is
         * that matching a build is not the same as executing correctly inside
         * it. Version, timestamp, image size and build signature together say
         * "these are the binaries the offsets were derived from"; they say
         * nothing about whether this machine's Office is configured, patched or
         * hooked in some way that changes what those offsets do.
         *
         * So the first process to see a given set of module identities runs the
         * deterministic self-test, and the verdict is remembered against those
         * identities. Later starts skip the proof - never the structural checks,
         * which have already run above and run again on every use.
         */
        for (const ModuleIdentity& module : candidate.modules) {
            if (!module.strong()) {
                std::wostringstream out;
                out << L"exact profile refused: " << module.name
                    << L" carries no build signature, so its identity is not strong enough to "
                       L"authorise private calls";
                return Portable(out.str());
            }
        }

        if (LookUpValidation(candidate) != CachedVerdict::Passed) {
            if (!selfTest) {
                return Portable(L"an unproven exact profile requires a self-test and none was "
                                L"given");
            }
            std::wstring failure;
            if (!selfTest(candidate, failure)) {
                // Not written to the store: a failure may be this host having a
                // bad moment, and condemning a build permanently on one run
                // would be worse than testing it again next time.
                return Portable(L"exact profile failed its first-use self-test: " + failure);
            }
            RecordValidationPassed(candidate);
        }

        Resolution resolution;
        resolution.decision = Decision::Native;
        resolution.profile = candidate;
        resolution.reason = L"exact profile: " + candidate.Describe();
        return resolution;
    }

    /*
     * Step 3: ask the architecture to resolve one structurally. It must refuse
     * on ambiguity; the framework cannot check that for it, which is why a
     * resolved profile is then required to pass a self-test that an exact one is
     * not.
     */
    NativeProfile resolved;
    std::wstring why;
    if (!resolver.TryResolve(world, resolved, why)) {
        std::wostringstream out;
        out << L"no exact profile matched and structural resolution declined: " << why;
        const std::wstring detail = rejected.str();
        if (!detail.empty()) {
            out << L"\nexact profiles considered:" << detail;
        }
        return Portable(out.str());
    }

    if (!resolved.valid() || resolved.architecture() != Current()) {
        return Portable(L"structural resolution produced a profile for the wrong architecture");
    }
    if (!resolved.facts->Revalidate()) {
        return Portable(L"structurally resolved profile did not revalidate");
    }

    /*
     * Step 4: prove it does the job. Structural validation says a candidate is
     * shaped right, never that it is the function meant, and the only thing that
     * can tell those apart is doing the work and checking the result.
     */
    if (!selfTest) {
        return Portable(L"a structurally resolved profile requires a self-test and none was given");
    }
    std::wstring failure;
    if (LookUpValidation(resolved) != CachedVerdict::Passed) {
        if (!selfTest(resolved, failure)) {
            return Portable(L"structurally resolved profile failed its self-test: " + failure);
        }
        RecordValidationPassed(resolved);
    }

    Resolution resolution;
    resolution.decision = Decision::Native;
    resolution.profile = resolved;
    resolution.reason = L"resolved and self-tested: " + resolved.Describe();
    return resolution;
}

namespace {

/**
 * The process-wide answer.
 *
 * Held as a pointer rather than a flag so that "not yet computed" and "computed,
 * and the answer was portable" are different states. The second must not cause a
 * re-resolve: a failed self-test means unvalidated code was driven into Office
 * once, and doing it again on the next call would turn one careful experiment
 * into a loop.
 */
std::unique_ptr<Resolution>& Slot() noexcept {
    static std::unique_ptr<Resolution> resolution;
    return resolution;
}

} // namespace

const Resolution&
ResolutionForThisProcess(const ArchitectureResolver& resolver,
                         const std::function<bool(const NativeProfile&, std::wstring&)>& selfTest) {
    std::unique_ptr<Resolution>& slot = Slot();
    if (!slot) {
        slot = std::make_unique<Resolution>(Resolve(resolver, selfTest));
    }
    return *slot;
}

void ForgetResolutionForTesting() noexcept {
    Slot().reset();
}

} // namespace bb::native
