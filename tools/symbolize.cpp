#include <dbghelp.h>
#include <filesystem>
#include <iostream>
#include <string>
#include <windows.h>

int wmain(int argc, wchar_t** argv) {
    if (argc < 4) {
        std::cerr << "bb_symbols module.dll cache rva...\n";
        return 2;
    }
    LoadLibraryExW(L"C:\\Program Files (x86)\\Windows Kits\\10\\App Certification Kit\\symsrv.dll",
                   nullptr,
                   LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32);
    auto process = GetCurrentProcess();
    std::filesystem::create_directories(argv[2]);
    auto path = L"srv*" + std::wstring(argv[2]) + L"*https://msdl.microsoft.com/download/symbols";
    SymSetOptions(SYMOPT_UNDNAME | SYMOPT_DEFERRED_LOADS | SYMOPT_FAIL_CRITICAL_ERRORS |
                  SYMOPT_NO_PROMPTS);
    if (!SymInitializeW(process, path.c_str(), FALSE)) {
        std::cerr << "SymInitialize " << GetLastError() << '\n';
        return 1;
    }
    auto base = SymLoadModuleExW(process, nullptr, argv[1], nullptr, 0x180000000, 0, nullptr, 0);
    if (!base) {
        std::cerr << "SymLoadModule " << GetLastError() << '\n';
        return 1;
    }
    for (int i = 3; i < argc; i++) {
        auto rva = std::stoull(argv[i], nullptr, 16);
        alignas(SYMBOL_INFOW) BYTE storage[sizeof(SYMBOL_INFOW) + 2048 * sizeof(wchar_t)]{};
        auto s = (SYMBOL_INFOW*)storage;
        s->SizeOfStruct = sizeof(*s);
        s->MaxNameLen = 2048;
        DWORD64 displacement = 0;
        if (SymFromAddrW(process, base + rva, &displacement, s)) {
            std::wcout << L"RVA 0x" << std::hex << rva << L" " << s->Name << L"+0x" << displacement
                       << L" symbol RVA=0x" << s->Address - base << '\n';
        } else {
            std::wcout << L"RVA 0x" << std::hex << rva << L" unavailable error=" << std::dec
                       << GetLastError() << '\n';
        }
    }
    IMAGEHLP_MODULEW64 m{};
    m.SizeOfStruct = sizeof(m);
    if (SymGetModuleInfoW64(process, base, &m)) {
        std::wcout << L"SymType=" << m.SymType << L" PDB=" << m.LoadedPdbName << '\n';
    }
    SymCleanup(process);
    return 0;
}
