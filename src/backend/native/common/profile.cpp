/**
 * @file profile.cpp
 * Profile matching. See the header for why the facts are opaque.
 */

#include "profile.hpp"

#include <sstream>

namespace bb::native {

const char* Describe(Provenance provenance) noexcept {
    switch (provenance) {
    case Provenance::ExactCompiled:
        return "exact, compiled in";
    case Provenance::Resolved:
        return "resolved at runtime";
    }
    return "unknown provenance";
}

std::wstring NativeProfile::Describe() const {
    std::wostringstream out;
    out << L"[" << Name(architecture()) << L", ";
    const char* how = bb::native::Describe(provenance);
    while (*how) {
        out << static_cast<wchar_t>(*how++);
    }
    out << L"] ";
    for (std::size_t index = 0; index < modules.size(); ++index) {
        if (index != 0) {
            out << L", ";
        }
        out << modules[index].Describe();
    }
    if (facts) {
        out << L" :: " << facts->Describe();
    }
    return out.str();
}

bool ProfileMatchesLoadedModules(const NativeProfile& profile, std::wstring& mismatch) {
    mismatch.clear();
    if (!profile.valid()) {
        mismatch = L"the profile carries no facts";
        return false;
    }

    /*
     * The architecture is checked first and separately, because it is the one
     * mismatch whose consequence is not "acceleration is unavailable" but
     * "addresses into the wrong binary". It cannot be allowed to be one failure
     * among several in a loop.
     */
    if (profile.architecture() != Current()) {
        std::wostringstream out;
        out << L"profile is for " << Name(profile.architecture()) << L", this build is "
            << Name(Current());
        mismatch = out.str();
        return false;
    }

    for (const ModuleIdentity& expected : profile.modules) {
        const ModuleIdentity actual = IdentifyLoadedModule(expected.name.c_str());
        if (!actual.valid()) {
            std::wostringstream out;
            out << expected.name << L" is not loaded, or could not be identified";
            mismatch = out.str();
            return false;
        }
        if (actual != expected) {
            // Both sides are reported. "oart.dll changed" is not actionable;
            // "expected ts=X size=Y, found ts=Z size=W" tells whoever reads it
            // that Office was updated and which build to derive against next.
            std::wostringstream out;
            out << L"expected " << expected.Describe() << L", found " << actual.Describe();
            mismatch = out.str();
            return false;
        }
    }
    return true;
}

} // namespace bb::native
