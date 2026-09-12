#pragma once
/**
 * @file resolver.hpp
 * Decides whether a native backend may run, and refuses by default.
 *
 * ## The decision, in order
 *
 * ```
 *   architecture matches the build?            no -> Portable
 *   required modules loaded and identified?    no -> Portable
 *   an exact compiled profile matches?         yes -> validate -> Native
 *   an architecture resolver offers one?       yes -> validate -> self-test -> Native
 *                                              otherwise    -> Portable
 * ```
 *
 * Every arrow that is not "yes" ends at Portable. That is the whole design: a
 * future Office update should cost acceleration, never a crash, so anything this
 * resolver is unsure about resolves to the backend that calls only documented
 * APIs.
 *
 * ## Exact profiles outrank resolved ones
 *
 * A compiled profile was derived offline against that exact build by someone
 * reading a disassembler. A resolved one was inferred at runtime from structural
 * evidence. Both are validated the same way, but they are not equally trusted,
 * and an exact match is taken in preference without consulting the resolver at
 * all.
 *
 * ## Resolved profiles must earn it twice
 *
 * Structural validation proves a candidate is *shaped* right. It cannot prove it
 * is the function meant. So a resolved profile additionally requires a
 * deterministic self-test - apply a known image to a disposable Shape and check
 * the result - before `BB_CAP_NATIVE_BACKEND` is reported. A profile whose
 * self-test fails is remembered as failed and not retried in this process: the
 * second attempt would fail the same way, and repeatedly driving unvalidated
 * code into Office to watch it not work is not diagnostics.
 *
 * ## What the framework does not know
 *
 * How to find anything. Discovery is architecture-specific and is supplied by
 * whichever backend is compiled - see `ArchitectureResolver`. The x86 probing
 * that motivated this design found x64's central structural invariant absent, so
 * a framework with discovery built into it would have been a framework only one
 * architecture could use.
 */

#include "profile.hpp"
#include "structural_validation.hpp"
#include <functional>
#include <memory>
#include <string>

namespace bb::native {

/// What the resolver concluded, and why.
enum class Decision {
    /// A validated profile is available; the native backend may run.
    Native,
    /// Nothing usable was found. The portable backend runs. This is not an error.
    Portable,
};

struct Resolution {
    Decision decision = Decision::Portable;
    NativeProfile profile;
    /// Always populated, including on success, so a diagnostic can say what matched.
    std::wstring reason;

    bool native() const noexcept {
        return decision == Decision::Native && profile.valid();
    }
};

/**
 * What an architecture supplies to the framework.
 *
 * Only the architecture knows what its Office internals look like. The framework
 * supplies the policy, the module world, the validation primitives and the
 * refusal-by-default.
 */
class ArchitectureResolver {
  public:
    virtual ~ArchitectureResolver() = default;

    /// The modules this architecture's backend needs, by name.
    virtual std::vector<const wchar_t*> RequiredModules() const = 0;

    /**
     * The profiles compiled in for exact builds, newest first.
     *
     * Returned rather than stored so that an architecture with none - which is
     * the honest state for a new one - simply returns an empty list.
     */
    virtual std::vector<NativeProfile> ExactProfiles() const = 0;

    /**
     * Attempts to resolve a profile structurally for a build with no exact match.
     *
     * Must refuse on any ambiguity. Returning a profile is a claim that exactly
     * one candidate was found for every fact and that each passed structural
     * validation; the framework then requires a self-test on top.
     *
     * An architecture that has not implemented adaptive resolution returns false,
     * which is a complete and correct answer.
     */
    virtual bool
    TryResolve(const ModuleWorld& world, NativeProfile& resolved, std::wstring& reason) const = 0;
};

/**
 * Runs the decision above.
 *
 * @p selfTest is invoked only for a resolved profile, never for an exact one,
 * and only after structural validation has passed. It must be deterministic and
 * must use disposable objects; the framework does not know how to make those.
 */
Resolution Resolve(const ArchitectureResolver& resolver,
                   const std::function<bool(const NativeProfile&, std::wstring&)>& selfTest);

/**
 * The resolution for this process, computed once.
 *
 * Cached because resolution reads a good deal of memory and the answer cannot
 * change while the modules do not: a module cannot be swapped under a running
 * process without unloading it, and the identities are re-checked on every use
 * anyway.
 *
 * A profile whose self-test failed is remembered as failed here.
 */
const Resolution&
ResolutionForThisProcess(const ArchitectureResolver& resolver,
                         const std::function<bool(const NativeProfile&, std::wstring&)>& selfTest);

/// Forgets the cached resolution. Tests only; nothing in the shipping path calls it.
void ForgetResolutionForTesting() noexcept;

} // namespace bb::native
