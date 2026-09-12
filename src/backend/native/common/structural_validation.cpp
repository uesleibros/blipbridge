/**
 * @file structural_validation.cpp
 * Implementation of the candidate checks. See the header for the rule they exist
 * to enforce.
 */

#include "structural_validation.hpp"

#include "memory_safety.hpp"
#include <psapi.h>
#include <sstream>

namespace bb::native {
namespace {

/**
 * How far a vtable pointer may be from pointer alignment: not at all.
 *
 * Both architectures align vtables and function entry points to at least the
 * pointer size. A misaligned candidate is not a marginal case to allow through -
 * it is a strong sign the read landed mid-structure.
 */
constexpr std::uintptr_t kPointerAlignmentMask = sizeof(void*) - 1;

} // namespace

const char* Describe(ValidationFailure failure) noexcept {
    switch (failure) {
    case ValidationFailure::None:
        return "valid";
    case ValidationFailure::NotInAnyKnownModule:
        return "the address is not inside any Office module this backend knows";
    case ValidationFailure::NotInExpectedModule:
        return "the address is in a different module from the one expected";
    case ValidationFailure::NotExecutable:
        return "the address is not in an executable section, so it is not code";
    case ValidationFailure::Misaligned:
        return "the address is not pointer-aligned";
    case ValidationFailure::VtableUnreadable:
        return "the vtable is not readable for the number of slots required";
    case ValidationFailure::VtableSlotNotCode:
        return "a vtable slot does not point at code in a known module";
    case ValidationFailure::VtableTooShort:
        return "the vtable has fewer usable slots than required";
    case ValidationFailure::Ambiguous:
        return "more than one candidate matched and none could be distinguished";
    case ValidationFailure::NoCandidate:
        return "nothing matched";
    }
    return "refused";
}

bool ModuleWorld::Add(const wchar_t* name) {
    if (!name) {
        return false;
    }
    // GetModuleHandleW, never LoadLibrary. Discovering what is loaded must not
    // change what is loaded - GFX in particular is delay-loaded, and pulling it
    // in as a side effect of a capability query would be a surprise.
    HMODULE module = GetModuleHandleW(name);
    if (!module) {
        return false;
    }
    MODULEINFO information{};
    if (!GetModuleInformation(GetCurrentProcess(), module, &information, sizeof(information))) {
        return false;
    }
    modules_.push_back(ModuleRange{
        name, reinterpret_cast<std::uintptr_t>(information.lpBaseOfDll), information.SizeOfImage});
    return true;
}

const ModuleRange* ModuleWorld::Find(std::uintptr_t address) const noexcept {
    for (const ModuleRange& module : modules_) {
        if (module.contains(address)) {
            return &module;
        }
    }
    return nullptr;
}

const ModuleRange* ModuleWorld::ByName(const wchar_t* name) const noexcept {
    if (!name) {
        return nullptr;
    }
    for (const ModuleRange& module : modules_) {
        if (_wcsicmp(module.name.c_str(), name) == 0) {
            return &module;
        }
    }
    return nullptr;
}

std::wstring ModuleWorld::Describe(std::uintptr_t address) const {
    std::wostringstream out;
    if (const ModuleRange* module = Find(address)) {
        out << module->name << L"+0x" << std::hex << (address - module->base) << std::dec;
    } else {
        out << L"0x" << std::hex << address << std::dec << L" (outside every known module)";
    }
    return out.str();
}

Validation
ValidateCodeAddress(const ModuleWorld& world, std::uintptr_t address, const ModuleRange* expected) {
    if (address == 0) {
        return Validation::Fail(ValidationFailure::NoCandidate);
    }

    /*
     * Module membership is asked first, before alignment. Both would refuse a
     * wild pointer, but they say different things: "not in any Office module"
     * identifies a pointer that has no business being a candidate at all, while
     * "misaligned" suggests a read that landed mid-structure inside a module
     * that was otherwise right. The second is a refinement of the first, and
     * reporting it for a stack address would send a reader looking in the wrong
     * place.
     */
    const ModuleRange* actual = world.Find(address);
    if (!actual) {
        return Validation::Fail(ValidationFailure::NotInAnyKnownModule, world.Describe(address));
    }
    if (expected && actual != expected) {
        std::wostringstream out;
        out << world.Describe(address) << L", expected " << expected->name;
        return Validation::Fail(ValidationFailure::NotInExpectedModule, out.str());
    }
    if ((address & kPointerAlignmentMask) != 0) {
        // Function entry points are aligned on both architectures, so this
        // catches a read that landed mid-pointer inside the right module.
        return Validation::Fail(ValidationFailure::Misaligned, world.Describe(address));
    }
    if (!IsExecutable(reinterpret_cast<const void*>(address))) {
        return Validation::Fail(ValidationFailure::NotExecutable, world.Describe(address));
    }
    return Validation::Pass();
}

Validation ValidateVtable(const ModuleWorld& world,
                          std::uintptr_t vtable,
                          std::size_t minimumSlots,
                          const ModuleRange* expected) {
    if (vtable == 0) {
        return Validation::Fail(ValidationFailure::NoCandidate);
    }
    if ((vtable & kPointerAlignmentMask) != 0) {
        return Validation::Fail(ValidationFailure::Misaligned, world.Describe(vtable));
    }
    if (minimumSlots == 0) {
        minimumSlots = 1;
    }
    if (!IsReadable(reinterpret_cast<const void*>(vtable), sizeof(void*) * minimumSlots)) {
        return Validation::Fail(ValidationFailure::VtableUnreadable, world.Describe(vtable));
    }
    if (expected && !expected->contains(vtable)) {
        std::wostringstream out;
        out << world.Describe(vtable) << L", expected " << expected->name;
        return Validation::Fail(ValidationFailure::NotInExpectedModule, out.str());
    }

    /*
     * Every slot, not just the first. An object of the wrong type still has a
     * first word and that word still points somewhere; requiring the whole table
     * to be code in modules this backend knows about is what turns "it looks
     * like a pointer" into "this is a vtable".
     */
    for (std::size_t slot = 0; slot < minimumSlots; ++slot) {
        const std::uintptr_t target =
            ReadPointer(reinterpret_cast<const void*>(vtable), sizeof(void*) * slot);
        if (target == 0) {
            std::wostringstream out;
            out << L"slot " << slot << L" is null";
            return Validation::Fail(ValidationFailure::VtableTooShort, out.str());
        }
        const Validation slotCheck = ValidateCodeAddress(world, target, nullptr);
        if (!slotCheck.ok) {
            std::wostringstream out;
            out << L"slot " << slot << L": " << slotCheck.detail;
            return Validation::Fail(ValidationFailure::VtableSlotNotCode, out.str());
        }
    }
    return Validation::Pass();
}

Validation RequireExactlyOne(const std::vector<std::uintptr_t>& candidates,
                             std::uintptr_t& chosen) {
    chosen = 0;
    if (candidates.empty()) {
        return Validation::Fail(ValidationFailure::NoCandidate);
    }
    if (candidates.size() > 1) {
        std::wostringstream out;
        out << candidates.size() << L" candidates";
        return Validation::Fail(ValidationFailure::Ambiguous, out.str());
    }
    chosen = candidates.front();
    return Validation::Pass();
}

} // namespace bb::native
