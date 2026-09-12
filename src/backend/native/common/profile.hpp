#pragma once
/**
 * @file profile.hpp
 * What a native backend needs to know about one exact Office build, and how it
 * is allowed to come by that knowledge.
 *
 * ## Why the facts are opaque here
 *
 * The obvious design is a struct with the offsets in it. That design was written
 * once already, for x64, and probing 32-bit Office showed why it must not be
 * generalised: the x64 profile's central fact is "the PPCORE wrapper's vtable is
 * N identity thunks with a consistent inner offset", and **32-bit Office has no
 * identity thunks at all**. Its wrapper methods are ordinary framed `__stdcall`
 * functions. A shared struct with an `innerThunkCount` field would be a field
 * one architecture could never fill and the other could never drop.
 *
 * So this header owns only what is genuinely invariant - identity, provenance,
 * confidence, and the rules about when a profile may be used - and each
 * architecture defines its own facts behind `ProfileFacts`. The framework
 * decides *whether* a profile applies; the architecture decides *what it says*.
 *
 * ## Provenance is part of the profile
 *
 * A profile that was derived by hand against a disassembler and one that was
 * resolved structurally at runtime are not equally trustworthy, and the
 * difference has to survive into the decision about what to enable. So it is
 * recorded rather than inferred.
 */

#include "module_identity.hpp"
#include <memory>
#include <string>
#include <vector>

namespace bb::native {

/// How a profile's facts were arrived at. Ordered by how much they are trusted.
enum class Provenance {
    /// Derived offline against this exact build and compiled in. Highest confidence.
    ExactCompiled,
    /// Resolved structurally at runtime and validated. Usable only after a self-test.
    Resolved,
};

const char* Describe(Provenance provenance) noexcept;

/**
 * The architecture-specific facts, opaque to the framework.
 *
 * Each native backend derives its own type from this. Nothing here inspects it;
 * the framework's whole job is to decide whether it may be handed over.
 */
class ProfileFacts {
  public:
    virtual ~ProfileFacts() = default;

    /// A one-line summary for diagnostics, in the architecture's own terms.
    virtual std::wstring Describe() const = 0;

    /**
     * Re-checks the facts against the live process.
     *
     * Called before the profile is used, every time it is fetched from a cache,
     * and never assumed to have been done. A profile is a set of claims about
     * memory; memory is what should be asked.
     */
    virtual bool Revalidate() const = 0;
};

/// One Office build a native backend can work against.
struct NativeProfile {
    /// The modules this profile's facts were derived from, all of them.
    std::vector<ModuleIdentity> modules;
    Provenance provenance = Provenance::ExactCompiled;
    std::shared_ptr<const ProfileFacts> facts;

    bool valid() const noexcept {
        return facts != nullptr && !modules.empty();
    }

    Architecture architecture() const noexcept {
        return modules.empty() ? Architecture::Unknown : modules.front().architecture;
    }

    std::wstring Describe() const;
};

/**
 * Whether @p profile describes the modules currently loaded.
 *
 * Every module the profile names must be present and identical - same
 * architecture, version, timestamp and image size. One mismatch refuses the
 * whole profile rather than the module: a profile is a set of offsets derived
 * together, and using the half that still matches would be using offsets into a
 * binary that has changed.
 */
bool ProfileMatchesLoadedModules(const NativeProfile& profile, std::wstring& mismatch);

} // namespace bb::native
