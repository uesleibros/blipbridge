/**
 * @file structure_probe.cpp
 * Read-only structural reconnaissance of the objects behind `Shape.Fill`.
 *
 * ## What this is for
 *
 * The accelerated backend on x64 reaches an OART receiver by walking from the
 * PPCORE wrapper that `Shape.Fill` hands back. None of that walk transfers to
 * 32-bit Office: different module, different layout, different thunk encoding,
 * different pointer size. Re-deriving it has to start from evidence, and this is
 * the instrument that gathers it.
 *
 * ## What it will not do
 *
 * **It never calls anything.** Every access is a guarded read of committed
 * memory. It does not call a vtable slot, does not construct an Office object,
 * does not touch a document. A probe that "tries it and sees" would be
 * establishing facts by crashing PowerPoint, which is not a method.
 *
 * It also draws no conclusions. It reports addresses, module attributions and
 * raw bytes; deciding what a thunk means is a human reading the output against a
 * disassembler. Hypotheses written into the probe would come back out of it
 * looking like findings.
 *
 * ## Why it is architecture-neutral
 *
 * The same source runs on x64 and x86, so the two can be compared directly and
 * any difference in the output is a difference in Office rather than in the
 * instrument. Pointer size comes from `sizeof(void*)` and nothing here assumes
 * either.
 *
 * Research only. Nothing in the shipping path calls this.
 */

#include "../../src/backend/native/common/architecture.hpp"
#include "../../src/backend/native/common/module_identity.hpp"
#include "../experiment_api.hpp"
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstdint>
#include <cstring>
#include <iomanip>
#include <sstream>
#include <string>
#include <vector>

// MinGW requires the Windows base types before the Automation declarations.
#include <windows.h>

#include <psapi.h>

namespace {

/// True when every page spanning [address, address+size) is committed and readable.
bool IsReadable(const void* address, std::size_t size) {
    if (!address || size == 0) {
        return false;
    }
    const auto* cursor = static_cast<const std::uint8_t*>(address);
    const std::uint8_t* end = cursor + size;
    while (cursor < end) {
        MEMORY_BASIC_INFORMATION information{};
        if (VirtualQuery(cursor, &information, sizeof(information)) != sizeof(information)) {
            return false;
        }
        if (information.State != MEM_COMMIT) {
            return false;
        }
        constexpr DWORD kNoRead = PAGE_NOACCESS | PAGE_GUARD;
        if ((information.Protect & kNoRead) != 0 || information.Protect == 0) {
            return false;
        }
        cursor = static_cast<const std::uint8_t*>(information.BaseAddress) + information.RegionSize;
    }
    return true;
}

/// One loaded module, for attributing an address to `name+0xRVA`.
struct LoadedModule {
    std::wstring name;
    std::uintptr_t base = 0;
    std::size_t size = 0;
};

std::vector<LoadedModule> LoadedModules() {
    std::vector<LoadedModule> modules;
    HMODULE handles[512];
    DWORD needed = 0;
    if (!EnumProcessModules(GetCurrentProcess(), handles, sizeof(handles), &needed)) {
        return modules;
    }
    const DWORD count = needed / sizeof(HMODULE);
    for (DWORD index = 0; index < count && index < 512; ++index) {
        MODULEINFO information{};
        wchar_t name[MAX_PATH] = {};
        if (!GetModuleInformation(
                GetCurrentProcess(), handles[index], &information, sizeof(information))) {
            continue;
        }
        GetModuleBaseNameW(GetCurrentProcess(), handles[index], name, MAX_PATH);
        modules.push_back({name,
                           reinterpret_cast<std::uintptr_t>(information.lpBaseOfDll),
                           information.SizeOfImage});
    }
    return modules;
}

const LoadedModule* ModuleFor(const std::vector<LoadedModule>& modules, std::uintptr_t address) {
    for (const LoadedModule& module : modules) {
        if (address >= module.base && address < module.base + module.size) {
            return &module;
        }
    }
    return nullptr;
}

/// `oart.dll+0x1A2B3C`, or a bare address when it belongs to no loaded module.
std::wstring Describe(const std::vector<LoadedModule>& modules, std::uintptr_t address) {
    std::wostringstream out;
    const LoadedModule* module = ModuleFor(modules, address);
    if (module) {
        out << module->name << L"+0x" << std::hex << (address - module->base) << std::dec;
    } else {
        out << L"0x" << std::hex << address << std::dec << L" (no module)";
    }
    return out.str();
}

/// Whether an address is inside an executable section - a vtable slot must be.
bool IsExecutable(const void* address) {
    MEMORY_BASIC_INFORMATION information{};
    if (VirtualQuery(address, &information, sizeof(information)) != sizeof(information)) {
        return false;
    }
    if (information.State != MEM_COMMIT) {
        return false;
    }
    constexpr DWORD kExecute =
        PAGE_EXECUTE | PAGE_EXECUTE_READ | PAGE_EXECUTE_READWRITE | PAGE_EXECUTE_WRITECOPY;
    return (information.Protect & kExecute) != 0;
}

std::wstring HexBytes(const std::uint8_t* bytes, std::size_t count) {
    std::wostringstream out;
    out << std::hex << std::setfill(L'0');
    for (std::size_t index = 0; index < count; ++index) {
        out << std::setw(2) << static_cast<unsigned>(bytes[index]) << L' ';
    }
    return out.str();
}

std::uintptr_t ReadPointer(const void* base, std::size_t offset) {
    std::uintptr_t value = 0;
    const auto* at = static_cast<const std::uint8_t*>(base) + offset;
    if (!IsReadable(at, sizeof(value))) {
        return 0;
    }
    std::memcpy(&value, at, sizeof(value));
    return value;
}

} // namespace

std::wstring inspectFillStructure(IDispatch* fill, long slotCount, long byteCount) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing Fill object");
    }
    if (slotCount <= 0 || slotCount > 128) {
        slotCount = 40;
    }
    if (byteCount <= 0 || byteCount > 64) {
        byteCount = 16;
    }

    const std::vector<LoadedModule> modules = LoadedModules();
    std::wostringstream out;
    out << L"pointerSize=" << sizeof(void*) << L";";

    const auto object = reinterpret_cast<const std::uint8_t*>(fill);
    if (!IsReadable(object, sizeof(void*))) {
        throw bb::Error(E_INVALIDARG, "The Fill pointer does not address committed memory");
    }

    const std::uintptr_t vtable = ReadPointer(object, 0);
    out << L"object=" << Describe(modules, reinterpret_cast<std::uintptr_t>(object)) << L";";
    out << L"vtable=" << Describe(modules, vtable) << L";";

    /*
     * The object's own leading words. On x64 the wrapper keeps the inner object
     * at a fixed offset and the walk reads it from there; whether 32-bit Office
     * does anything similar is exactly what this is for. Reported raw, with each
     * word attributed to a module when it looks like a pointer, so a reader can
     * see which of them even could be one.
     */
    out << L"\r\nobject words:\r\n";
    for (std::size_t offset = 0; offset < 0x80; offset += sizeof(void*)) {
        const std::uintptr_t word = ReadPointer(object, offset);
        if (word == 0) {
            continue;
        }
        const LoadedModule* module = ModuleFor(modules, word);
        const bool pointsToCommitted = IsReadable(reinterpret_cast<const void*>(word), 4);
        out << L"  +0x" << std::hex << offset << std::dec << L"  0x" << std::hex << word
            << std::dec;
        if (module) {
            out << L"  " << Describe(modules, word);
        } else if (pointsToCommitted) {
            // A heap pointer: read its own first word, which for a COM-ish
            // object would be its vtable and is the thread worth pulling.
            const std::uintptr_t inner = ReadPointer(reinterpret_cast<const void*>(word), 0);
            out << L"  heap, [0]=" << Describe(modules, inner);
        }
        out << L"\r\n";
    }

    if (!IsReadable(reinterpret_cast<const void*>(vtable), sizeof(void*) * slotCount)) {
        out << L"vtableReadable=0;";
        return out.str();
    }

    out << L"\r\nvtable slots:\r\n";
    for (long slot = 0; slot < slotCount; ++slot) {
        const std::uintptr_t target =
            ReadPointer(reinterpret_cast<const void*>(vtable), sizeof(void*) * slot);
        if (target == 0) {
            continue;
        }
        out << L"  [" << std::setw(3) << slot << L"] +0x" << std::hex << (sizeof(void*) * slot)
            << std::dec << L"  " << Describe(modules, target);
        if (!IsExecutable(reinterpret_cast<const void*>(target))) {
            out << L"  NOT-EXECUTABLE";
        } else if (IsReadable(reinterpret_cast<const void*>(target),
                              static_cast<std::size_t>(byteCount))) {
            out << L"  "
                << HexBytes(reinterpret_cast<const std::uint8_t*>(target),
                            static_cast<std::size_t>(byteCount));
        }
        out << L"\r\n";
    }
    return out.str();
}

/**
 * Reports the identities of the Office modules, through the compatibility
 * framework rather than around it.
 *
 * Two things at once: it produces the data an exact profile needs, and it
 * exercises the framework's identity reading inside the host it will have to
 * work in. A profile transcribed from a different tool's output would be a
 * profile nothing had checked the reader against.
 */
std::wstring inspectModuleIdentities() {
    std::wostringstream out;
    out << L"build architecture: " << bb::native::Name(bb::native::Current()) << L"\r\n";
    out << L"pointer size      : " << sizeof(void*) << L"\r\n\r\n";

    for (const wchar_t* name : {L"POWERPNT.EXE",
                                L"ppcore.dll",
                                L"oart.dll",
                                L"gfx.dll",
                                L"mso.dll",
                                L"mso20win32client.dll"}) {
        const bb::native::ModuleIdentity identity = bb::native::IdentifyLoadedModule(name);
        if (!identity.valid()) {
            out << L"  " << name << L": not loaded\r\n";
            continue;
        }
        out << L"  " << identity.Describe() << L"\r\n";
        out << L"      strong=" << (identity.strong() ? L"yes" : L"NO - cannot authorise")
            << L" arch=" << bb::native::Name(identity.architecture) << L"\r\n";
    }
    return out.str();
}
