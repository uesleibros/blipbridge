/**
 * @file native_framework_contract.cpp
 * The compatibility framework's contract, and above all its refusals.
 *
 * This runs off-host on both architectures and needs no Office. That is possible
 * because the framework deliberately knows nothing about Office: it decides
 * whether a profile may be used, never what one says, so its policy can be
 * driven with a fake architecture resolver and asserted exactly.
 *
 * ## What is worth testing here
 *
 * Almost all of it is the negative space. A resolver that returns Native when
 * everything is perfect is easy; the value is in the four ways it must return
 * Portable instead, and in the rule that a structurally resolved profile has to
 * pass a self-test that an exact one does not. Those are the behaviours that
 * stand between an Office update and a crash in someone's presentation, and they
 * are the ones nothing else would notice breaking.
 */

#include "../src/backend/native/common/architecture.hpp"
#include "../src/backend/native/common/memory_safety.hpp"
#include "../src/backend/native/common/module_identity.hpp"
#include "../src/backend/native/common/profile.hpp"
#include "../src/backend/native/common/resolver.hpp"
#include "../src/backend/native/common/structural_validation.hpp"
#include <cstdio>
#include <string>
#include <vector>

namespace {

int g_failures = 0;

void Check(bool condition, const std::string& what) {
    if (condition) {
        std::printf("  ok   %s\n", what.c_str());
        return;
    }
    std::printf("  FAIL %s\n", what.c_str());
    ++g_failures;
}

using namespace bb::native;

// --- architecture ---------------------------------------------------------

void ArchitectureChecks() {
    // The build knows what it is at compile time, and it must agree with the
    // pointer size. A mismatch here would mean Current() is lying, and every
    // profile decision rests on it.
    const Architecture current = Current();
    Check(current != Architecture::Unknown, "the build reports a known architecture");
    Check((current == Architecture::X64) == (sizeof(void*) == 8),
          "and it agrees with the pointer size");

    // A module this process has definitely loaded must read as this architecture.
    HMODULE self = GetModuleHandleW(nullptr);
    Check(ArchitectureOf(self) == current, "the running image reads as the same architecture");
    Check(MatchesCurrent(self), "and MatchesCurrent agrees");

    Check(ArchitectureOf(nullptr) == Architecture::Unknown,
          "a null module is Unknown, not a guess");
    Check(!MatchesCurrent(nullptr), "and never matches");

    // Not a module at all: a stack address. The header walk must refuse rather
    // than read whatever happens to be there.
    int onTheStack = 0;
    Check(ArchitectureOf(reinterpret_cast<HMODULE>(&onTheStack)) == Architecture::Unknown,
          "a non-module address is Unknown rather than misread");
}

// --- guarded reads ---------------------------------------------------------

void MemorySafetyChecks() {
    int value = 0x1234;
    Check(IsReadable(&value, sizeof(value)), "a live stack variable is readable");
    Check(!IsReadable(nullptr, 4), "null is not readable");
    Check(!IsReadable(&value, 0), "a zero-length read is refused rather than trivially allowed");

    // Address 0x10 is reserved and never committed. This is the case every
    // guarded read exists for: a plausible-looking small pointer.
    Check(!IsReadable(reinterpret_cast<const void*>(0x10), 4),
          "a low reserved address is not readable");

    // A size that wraps the address space must be refused, not truncated.
    Check(!IsReadable(&value, static_cast<std::size_t>(-1)), "a range that wraps is refused");

    Check(IsExecutable(reinterpret_cast<const void*>(&MemorySafetyChecks)),
          "a function's address is executable");
    Check(!IsExecutable(&value), "a stack variable is not executable");
    Check(!IsExecutable(nullptr), "null is not executable");

    Check(ReadPointer(nullptr, 0) == 0, "reading through null yields 0 rather than faulting");
    Check(ReadDword(reinterpret_cast<const void*>(0x10), 0) == 0,
          "reading unreadable memory yields 0");

    const std::uint8_t pattern[] = {0xDE, 0xAD, 0xBE, 0xEF};
    Check(BytesMatch(pattern, pattern, sizeof(pattern)), "bytes match themselves");
    const std::uint8_t other[] = {0xDE, 0xAD, 0xBE, 0x00};
    Check(!BytesMatch(pattern, other, sizeof(other)), "and differ when they differ");
    Check(!BytesMatch(reinterpret_cast<const void*>(0x10), pattern, sizeof(pattern)),
          "and a match against unreadable memory is false, not a fault");
}

// --- module identity -------------------------------------------------------

void ModuleIdentityChecks() {
    const ModuleIdentity kernel = IdentifyLoadedModule(L"kernel32.dll");
    Check(kernel.valid(), "kernel32 is identifiable");
    Check(kernel.architecture == Current(), "and reads as this architecture");
    Check(kernel.imageSize != 0, "with a non-zero image size");
    Check(!kernel.name.empty(), "and a name");

    // Identity is a value: asking twice must give the same answer, or it could
    // not be a cache key.
    Check(IdentifyLoadedModule(L"kernel32.dll") == kernel, "identity is stable across reads");

    // A different module must not compare equal, however similar.
    const ModuleIdentity ntdll = IdentifyLoadedModule(L"ntdll.dll");
    if (ntdll.valid()) {
        Check(ntdll != kernel, "two different modules are not the same identity");
    }

    Check(!IdentifyLoadedModule(L"this-module-does-not-exist.dll").valid(),
          "a module that is not loaded is not identified");
    Check(!IdentifyLoadedModule(nullptr).valid(), "and neither is a null name");

    // The timestamp and image size are what actually distinguish a rebuild, so a
    // profile that ignored them would accept the wrong binary.
    ModuleIdentity rebuilt = kernel;
    rebuilt.timestamp ^= 1u;
    Check(rebuilt != kernel, "a different PE timestamp is a different module");
    ModuleIdentity resized = kernel;
    resized.imageSize += 0x1000;
    Check(resized != kernel, "and so is a different image size");
}

// --- structural validation -------------------------------------------------

void StructuralValidationChecks() {
    ModuleWorld world;
    Check(world.Add(L"kernel32.dll"), "a loaded module can join the world");
    Check(!world.Add(L"this-module-does-not-exist.dll"), "an absent one cannot");
    Check(!world.Add(nullptr), "and neither can a null name");

    const ModuleRange* kernel = world.ByName(L"kernel32.dll");
    Check(kernel != nullptr, "and it can be found again by name");

    // A real export is code in that module.
    const auto real = reinterpret_cast<std::uintptr_t>(
        GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "GetProcAddress"));
    Check(real != 0, "a known export resolves");
    Check(ValidateCodeAddress(world, real, kernel).ok, "and validates as code in kernel32");

    // The same address refuses when a different module is demanded. This is the
    // check that stops a profile's offsets being read against the wrong binary.
    ModuleWorld twoModules;
    twoModules.Add(L"kernel32.dll");
    const bool haveNtdll = twoModules.Add(L"ntdll.dll");
    if (haveNtdll) {
        Check(!ValidateCodeAddress(twoModules, real, twoModules.ByName(L"ntdll.dll")).ok,
              "a kernel32 address is refused when ntdll was expected");
        Check(ValidateCodeAddress(twoModules, real, twoModules.ByName(L"ntdll.dll")).failure ==
                  ValidationFailure::NotInExpectedModule,
              "and the refusal names the reason");
    }

    int onTheStack = 0;
    const auto stack = reinterpret_cast<std::uintptr_t>(&onTheStack);
    Check(!ValidateCodeAddress(world, stack, nullptr).ok, "a stack address is not code");
    Check(ValidateCodeAddress(world, stack, nullptr).failure ==
              ValidationFailure::NotInAnyKnownModule,
          "and is refused for being outside every known module");

    Check(!ValidateCodeAddress(world, 0, nullptr).ok, "a null candidate is refused");
    Check(ValidateCodeAddress(world, real + 1, kernel).failure == ValidationFailure::Misaligned,
          "a misaligned candidate is refused as misaligned");

    // A fabricated vtable of real code pointers validates; one with a data
    // pointer in it does not. The second is the case that matters - an object of
    // the wrong type still has a first word.
    std::uintptr_t goodVtable[4] = {real, real, real, real};
    Check(ValidateVtable(world, reinterpret_cast<std::uintptr_t>(goodVtable), 4).ok,
          "a table of code pointers validates as a vtable");

    std::uintptr_t mixedVtable[4] = {real, real, stack, real};
    const Validation mixed =
        ValidateVtable(world, reinterpret_cast<std::uintptr_t>(mixedVtable), 4);
    Check(!mixed.ok, "a table with a data pointer in it does not");
    Check(mixed.failure == ValidationFailure::VtableSlotNotCode, "and says which check failed");

    std::uintptr_t shortVtable[4] = {real, real, 0, real};
    Check(ValidateVtable(world, reinterpret_cast<std::uintptr_t>(shortVtable), 4).failure ==
              ValidationFailure::VtableTooShort,
          "a null slot is refused rather than skipped");

    Check(!ValidateVtable(world, 0, 1).ok, "a null vtable is refused");
    Check(ValidateVtable(world, reinterpret_cast<std::uintptr_t>(goodVtable), 4, kernel).failure ==
              ValidationFailure::NotInExpectedModule,
          "a vtable outside the expected module is refused");

    // Ambiguity is a refusal, not a choice.
    std::uintptr_t chosen = 0;
    Check(RequireExactlyOne({}, chosen).failure == ValidationFailure::NoCandidate,
          "no candidates is NoCandidate");
    Check(RequireExactlyOne({real, real + 8}, chosen).failure == ValidationFailure::Ambiguous,
          "two candidates is Ambiguous, never the first one");
    Check(chosen == 0, "and nothing is chosen when it is ambiguous");
    Check(RequireExactlyOne({real}, chosen).ok && chosen == real, "exactly one is taken");
}

// --- the resolver's policy -------------------------------------------------

/// Facts that do as they are told, so the policy can be tested rather than Office.
class FakeFacts final : public ProfileFacts {
  public:
    explicit FakeFacts(bool revalidates) : revalidates_(revalidates) {}

    std::wstring Describe() const override {
        return L"fake facts";
    }

    bool Revalidate() const override {
        return revalidates_;
    }

  private:
    bool revalidates_;
};

NativeProfile MakeProfile(const wchar_t* moduleName, bool revalidates, Provenance provenance) {
    NativeProfile profile;
    profile.modules.push_back(IdentifyLoadedModule(moduleName));
    profile.provenance = provenance;
    profile.facts = std::make_shared<FakeFacts>(revalidates);
    return profile;
}

class FakeResolver final : public ArchitectureResolver {
  public:
    std::vector<const wchar_t*> required{L"kernel32.dll"};
    std::vector<NativeProfile> exact;
    bool offersResolved = false;
    NativeProfile resolved;

    std::vector<const wchar_t*> RequiredModules() const override {
        return required;
    }

    std::vector<NativeProfile> ExactProfiles() const override {
        return exact;
    }

    bool TryResolve(const ModuleWorld&, NativeProfile& out, std::wstring& reason) const override {
        if (!offersResolved) {
            reason = L"this architecture has no adaptive resolver";
            return false;
        }
        out = resolved;
        return true;
    }
};

void ResolverChecks() {
    const auto passes = [](const NativeProfile&, std::wstring&) { return true; };
    const auto fails = [](const NativeProfile&, std::wstring& why) {
        why = L"the applied image did not match";
        return false;
    };

    // A module that is not loaded means no native backend, and it is not an error.
    {
        FakeResolver resolver;
        resolver.required = {L"this-module-does-not-exist.dll"};
        const Resolution resolution = Resolve(resolver, passes);
        Check(!resolution.native(), "a missing required module resolves to portable");
        Check(resolution.reason.find(L"not loaded") != std::wstring::npos,
              "and says which module was missing");
    }

    // Nothing offered at all.
    {
        FakeResolver resolver;
        const Resolution resolution = Resolve(resolver, passes);
        Check(!resolution.native(), "no profiles at all resolves to portable");
    }

    // An exact profile that matches is taken, and is *not* self-tested: an
    // offline-derived profile for this exact build is the strongest evidence
    // there is, and re-proving it on every start would be cost for nothing.
    {
        FakeResolver resolver;
        resolver.exact.push_back(MakeProfile(L"kernel32.dll", true, Provenance::ExactCompiled));
        bool selfTested = false;
        const Resolution resolution = Resolve(resolver, [&](const NativeProfile&, std::wstring&) {
            selfTested = true;
            return true;
        });
        Check(resolution.native(), "a matching exact profile resolves to native");
        Check(!selfTested, "and is not put through the self-test");
    }

    // An exact profile whose module has changed must not be used.
    {
        FakeResolver resolver;
        NativeProfile stale = MakeProfile(L"kernel32.dll", true, Provenance::ExactCompiled);
        stale.modules.front().timestamp ^= 1u;
        resolver.exact.push_back(stale);
        const Resolution resolution = Resolve(resolver, passes);
        Check(!resolution.native(), "an exact profile for a different build is refused");
    }

    // Identity matched but the memory did not: refuse, and do not fall through
    // to adaptive resolution. Something is loaded that reports itself as a build
    // it does not match, which is worse than an unknown build.
    {
        FakeResolver resolver;
        resolver.exact.push_back(MakeProfile(L"kernel32.dll", false, Provenance::ExactCompiled));
        resolver.offersResolved = true;
        resolver.resolved = MakeProfile(L"kernel32.dll", true, Provenance::Resolved);
        const Resolution resolution = Resolve(resolver, passes);
        Check(!resolution.native(), "an exact profile that fails revalidation is refused");
    }

    // A resolved profile is taken only after a self-test.
    {
        FakeResolver resolver;
        resolver.offersResolved = true;
        resolver.resolved = MakeProfile(L"kernel32.dll", true, Provenance::Resolved);
        bool selfTested = false;
        const Resolution resolution = Resolve(resolver, [&](const NativeProfile&, std::wstring&) {
            selfTested = true;
            return true;
        });
        Check(resolution.native(), "a resolved profile that self-tests resolves to native");
        Check(selfTested, "and the self-test really ran");
    }

    // The self-test is what makes a resolved profile safe, so failing it is fatal.
    {
        FakeResolver resolver;
        resolver.offersResolved = true;
        resolver.resolved = MakeProfile(L"kernel32.dll", true, Provenance::Resolved);
        const Resolution resolution = Resolve(resolver, fails);
        Check(!resolution.native(), "a resolved profile that fails its self-test is refused");
        Check(resolution.reason.find(L"self-test") != std::wstring::npos, "and the reason says so");
    }

    // No self-test supplied is not permission to skip it.
    {
        FakeResolver resolver;
        resolver.offersResolved = true;
        resolver.resolved = MakeProfile(L"kernel32.dll", true, Provenance::Resolved);
        const Resolution resolution = Resolve(resolver, {});
        Check(!resolution.native(), "a resolved profile with no self-test available is refused");
    }

    // A profile for the other architecture must never be used, however well it
    // validates. This is the one mismatch whose cost is not lost acceleration.
    {
        FakeResolver resolver;
        NativeProfile foreign = MakeProfile(L"kernel32.dll", true, Provenance::ExactCompiled);
        foreign.modules.front().architecture =
            Current() == Architecture::X64 ? Architecture::X86 : Architecture::X64;
        resolver.exact.push_back(foreign);
        const Resolution resolution = Resolve(resolver, passes);
        Check(!resolution.native(), "a profile for the other architecture is refused");
    }

    // Every outcome carries a reason, including success. A diagnostic that says
    // only "portable" cannot be acted on.
    {
        FakeResolver resolver;
        Check(!Resolve(resolver, passes).reason.empty(), "a refusal always carries a reason");
        resolver.exact.push_back(MakeProfile(L"kernel32.dll", true, Provenance::ExactCompiled));
        Check(!Resolve(resolver, passes).reason.empty(), "and so does a success");
    }
}

void ProfileChecks() {
    NativeProfile empty;
    Check(!empty.valid(), "a profile with no facts is not valid");
    Check(empty.architecture() == Architecture::Unknown, "and has no architecture");

    std::wstring mismatch;
    Check(!ProfileMatchesLoadedModules(empty, mismatch), "an empty profile matches nothing");
    Check(!mismatch.empty(), "and says why");

    const NativeProfile good = MakeProfile(L"kernel32.dll", true, Provenance::ExactCompiled);
    Check(good.valid(), "a profile with facts and modules is valid");
    Check(ProfileMatchesLoadedModules(good, mismatch), "and matches the loaded module");
    Check(!good.Describe().empty(), "and describes itself");
}

} // namespace

int main() {
    std::printf("native framework contract\n");
    ArchitectureChecks();
    MemorySafetyChecks();
    ModuleIdentityChecks();
    StructuralValidationChecks();
    ProfileChecks();
    ResolverChecks();

    if (g_failures != 0) {
        std::printf("%d failure(s)\n", g_failures);
        return 1;
    }
    std::printf("all checks passed\n");
    return 0;
}
