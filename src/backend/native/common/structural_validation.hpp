#pragma once
/**
 * @file structural_validation.hpp
 * The checks that stand between "something matched" and "this may be called".
 *
 * ## The rule this file exists to enforce
 *
 * **A pattern match never enables the native backend.** Discovery and proof are
 * separate steps, and the gap between them is where a library like this one
 * either stays safe or corrupts a document. A byte pattern found at an address
 * says a sequence of bytes is there; it says nothing about whether that address
 * is a function, is in the module it should be in, or is the function meant.
 *
 * So everything a resolver proposes arrives here as a *candidate*, and a
 * candidate becomes usable only by passing checks that are about structure
 * rather than about bytes:
 *
 *  - it lies inside the module it is supposed to;
 *  - code lies in an executable section, data does not;
 *  - a vtable's slots all point at executable code in expected modules;
 *  - alignment is what the architecture requires;
 *  - pointers reached by walking stay inside modules that make sense.
 *
 * ## Ambiguity is a refusal
 *
 * If discovery leaves more than one candidate and nothing distinguishes them,
 * the answer is not "take the first". It is that the native backend stays off
 * and the portable one runs. A wrong address that happens to be executable is
 * indistinguishable from a right one until it is called, and calling it is the
 * thing being avoided.
 *
 * ## What is not here
 *
 * No architecture-specific knowledge. There are no x64 instruction shapes and no
 * x86 ones; this is the frame, and each architecture supplies its own
 * invariants to be checked within it. The x86 probing that motivated this
 * framework found that x64's structural signature - a vtable of identity thunks -
 * simply does not exist on x86, so a frame that assumed it would have been a
 * frame only one architecture could use.
 */

#include "architecture.hpp"
#include "module_identity.hpp"
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

// MinGW requires the Windows base types first.
#include <windows.h>

namespace bb::native {

/// A module the native backend is allowed to find things in.
struct ModuleRange {
    std::wstring name;
    std::uintptr_t base = 0;
    std::size_t size = 0;

    bool contains(std::uintptr_t address) const noexcept {
        return base != 0 && address >= base && address < base + size;
    }
};

/**
 * The modules a walk may legitimately land in, and nothing else.
 *
 * An address outside every one of them is not a candidate however plausible it
 * looks: it might be a third-party hook, a JIT page, or a stale pointer into
 * freed memory that happens to still be committed.
 */
class ModuleWorld {
  public:
    /// Adds a loaded module by name. Absent modules are simply not added.
    bool Add(const wchar_t* name);

    const ModuleRange* Find(std::uintptr_t address) const noexcept;
    const ModuleRange* ByName(const wchar_t* name) const noexcept;

    bool Contains(std::uintptr_t address) const noexcept {
        return Find(address) != nullptr;
    }

    /// `oart.dll+0x1A2B3C`, or the bare address when it belongs to nothing known.
    std::wstring Describe(std::uintptr_t address) const;

    const std::vector<ModuleRange>& modules() const noexcept {
        return modules_;
    }

  private:
    std::vector<ModuleRange> modules_;
};

/// Why a candidate was refused, in terms that name the check rather than the symptom.
enum class ValidationFailure {
    None,
    NotInAnyKnownModule,
    NotInExpectedModule,
    NotExecutable,
    Misaligned,
    VtableUnreadable,
    VtableSlotNotCode,
    VtableTooShort,
    Ambiguous,
    NoCandidate,
};

const char* Describe(ValidationFailure failure) noexcept;

/// The outcome of validating one candidate: a verdict plus why.
struct Validation {
    bool ok = false;
    ValidationFailure failure = ValidationFailure::NoCandidate;
    std::wstring detail;

    static Validation Pass() {
        return Validation{true, ValidationFailure::None, {}};
    }

    static Validation Fail(ValidationFailure failure, std::wstring detail = {}) {
        return Validation{false, failure, std::move(detail)};
    }
};

/**
 * A code address: inside @p expected, executable, and sanely aligned.
 *
 * @p expected may be null to mean "any module in the world", which is the right
 * looseness for a call target reached by walking and the wrong looseness for an
 * entry point a profile names.
 */
Validation
ValidateCodeAddress(const ModuleWorld& world, std::uintptr_t address, const ModuleRange* expected);

/**
 * A vtable: readable, at least @p minimumSlots long, every slot executable code
 * in a known module.
 *
 * This is the check that most often catches "the object is not what we think",
 * because an object of the wrong type still has a first word, and that word
 * still points somewhere. Requiring every slot to be code in a module the
 * backend knows about is a much stronger statement than requiring the first one
 * to be.
 */
Validation ValidateVtable(const ModuleWorld& world,
                          std::uintptr_t vtable,
                          std::size_t minimumSlots,
                          const ModuleRange* expected = nullptr);

/**
 * Reduces a candidate set to exactly one, or refuses.
 *
 * The whole point: "several things matched" is not a result. A caller that
 * wanted the first match could have written that, and would have been wrong the
 * first time two functions shared a prologue.
 */
Validation RequireExactlyOne(const std::vector<std::uintptr_t>& candidates, std::uintptr_t& chosen);

} // namespace bb::native
