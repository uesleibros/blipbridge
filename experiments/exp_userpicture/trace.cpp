// Isolated, reversible IAT instrumentation. Never enabled by normal Engine calls.
#include <atomic>
#include <blipbridge/dispatch.hpp>
#include <blipbridge/image_validation.hpp>
#include <filesystem>
#include <fstream>
#include <map>
#include <mutex>
#include <psapi.h>
#include <string>
#include <vector>
#include <windows.h>

namespace {
std::ofstream logFile;
std::mutex logMutex;
HANDLE tracked = INVALID_HANDLE_VALUE;
std::atomic<DWORD> traceThread{0};
bool memoryMode = false;
std::vector<BYTE> memoryBytes;
size_t memoryPosition = 0;
std::wstring memoryName;
std::map<HANDLE, std::wstring> cacheHandles;
unsigned long long openCalls = 0, readCalls = 0, writeCalls = 0;

bool isMemory(HANDLE h) {
    return GetCurrentThreadId() == traceThread && memoryMode && h == tracked &&
           h != INVALID_HANDLE_VALUE;
}

using OpenFn = decltype(&CreateFileW);
using ReadFn = decltype(&ReadFile);
using CloseFn = decltype(&CloseHandle);
OpenFn realOpen = CreateFileW;
ReadFn realRead = ReadFile;
CloseFn realClose = CloseHandle;

void stack(const char* event) noexcept {
    try {
        std::lock_guard lock(logMutex);
        logFile << "\nEVENT " << event << " thread=" << GetCurrentThreadId() << '\n';
        void* frames[48];
        USHORT n = CaptureStackBackTrace(1, 48, frames, nullptr);
        for (USHORT i = 0; i < n; i++) {
            HMODULE m = nullptr;
            wchar_t path[32768]{};
            GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                   GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                               (LPCWSTR)frames[i],
                               &m);
            GetModuleFileNameW(m, path, 32768);
            auto name = std::filesystem::path(path).filename().string();
            logFile << name << "+0x" << std::hex << ((uintptr_t)frames[i] - (uintptr_t)m)
                    << " absolute=0x" << (uintptr_t)frames[i] << std::dec << '\n';
        }
        logFile.flush();
    } catch (...) {
        OutputDebugStringW(L"BlipBridge trace logging failed\n");
    }
}

HANDLE WINAPI openHook(LPCWSTR name,
                       DWORD access,
                       DWORD share,
                       LPSECURITY_ATTRIBUTES sa,
                       DWORD disposition,
                       DWORD flags,
                       HANDLE templ) {
    if (GetCurrentThreadId() == traceThread) {
        ++openCalls;
    }
    bool match =
        GetCurrentThreadId() == traceThread && name &&
        (memoryMode ? memoryName == name : wcsstr(name, L"BB_TRACE_UNIQUE_TEXTURE_01") != nullptr);
    if (match) {
        stack("CreateFileW entry (unique texture)");
    }
    if (match && memoryMode) {
        if (access != GENERIC_READ || disposition != OPEN_EXISTING) {
            SetLastError(ERROR_ACCESS_DENIED);
            return INVALID_HANDLE_VALUE;
        }
        tracked = CreateEventW(nullptr, TRUE, TRUE, nullptr);
        if (!tracked) {
            tracked = INVALID_HANDLE_VALUE;
            return tracked;
        }
        memoryPosition = 0;
        logFile << "MEMORY open: event token, no file opened; bytes=" << memoryBytes.size() << '\n';
        SetLastError(ERROR_SUCCESS);
        return tracked;
    }
    HANDLE h = realOpen(name, access, share, sa, disposition, flags, templ);
    DWORD error = GetLastError();
    if (GetCurrentThreadId() == traceThread && name && wcsstr(name, L"Content.MSO") &&
        h != INVALID_HANDLE_VALUE) {
        try {
            cacheHandles[h] = name;
            stack("CreateFileW Office Content.MSO cache");
            logFile << "cacheHandle=" << h << " access=" << access
                    << " path=" << std::filesystem::path(name).string() << '\n';
        } catch (...) {
        }
    }
    if (match) {
        tracked = h;
        logFile << "CreateFileW result=" << h << " access=" << access << " flags=" << flags
                << " error=" << error << '\n';
        logFile.flush();
    }
    SetLastError(error);
    return h;
}

BOOL WINAPI readHook(HANDLE h, LPVOID p, DWORD n, LPDWORD got, LPOVERLAPPED ov) {
    if (GetCurrentThreadId() == traceThread) {
        ++readCalls;
        if (!cacheHandles.count(h)) {
            wchar_t path[32768]{};
            DWORD size = GetFinalPathNameByHandleW(h, path, 32768, FILE_NAME_NORMALIZED);
            if (size && size < 32768 && wcsstr(path, L"Content.MSO")) {
                try {
                    cacheHandles[h] = path;
                } catch (...) {
                }
            }
        }
    }
    if (GetCurrentThreadId() == traceThread && h == tracked) {
        stack("ReadFile unique texture");
        logFile << "requested=" << n << '\n';
    }
    if (GetCurrentThreadId() == traceThread && cacheHandles.count(h)) {
        stack("ReadFile Office cache");
        logFile << "cacheReadHandle=" << h << " bytes=" << n << '\n';
    }
    if (isMemory(h)) {
        size_t pos = ov ? ((uint64_t(ov->OffsetHigh) << 32) | ov->Offset) : memoryPosition;
        DWORD amount =
            pos < memoryBytes.size() ? (DWORD)std::min<size_t>(n, memoryBytes.size() - pos) : 0;
        if (amount) {
            memcpy(p, memoryBytes.data() + pos, amount);
        }
        if (got) {
            *got = amount;
        }
        if (ov) {
            ov->Internal = 0;
            ov->InternalHigh = amount;
            if (ov->hEvent) {
                SetEvent((HANDLE)((uintptr_t)ov->hEvent & ~uintptr_t(1)));
            }
        } else {
            memoryPosition += amount;
        }
        logFile << "MEMORY read offset=" << pos << " bytes=" << amount << '\n';
        SetLastError(ERROR_SUCCESS);
        return TRUE;
    }
    return realRead(h, p, n, got, ov);
}

BOOL WINAPI closeHook(HANDLE h) {
    if (GetCurrentThreadId() == traceThread) {
        if (h == tracked) {
            stack("CloseHandle unique texture");
            tracked = INVALID_HANDLE_VALUE;
        }
        cacheHandles.erase(h);
    }
    return realClose(h);
}

HRESULT WINAPI streamHook(HGLOBAL h, BOOL release, IStream** stream) {
    auto result = CreateStreamOnHGlobal(h, release, stream);
    if (GetCurrentThreadId() == traceThread) {
        stack("CreateStreamOnHGlobal");
        try {
            logFile << "stream=" << (SUCCEEDED(result) ? *stream : nullptr)
                    << " size=" << (h ? GlobalSize(h) : 0) << '\n';
        } catch (...) {
        }
    }
    return result;
}

HRESULT WINAPI
coCreateHook(REFCLSID cls, IUnknown* outer, DWORD context, REFIID iid, void** result) {
    if (GetCurrentThreadId() == traceThread) {
        wchar_t c[40]{}, i[40]{};
        StringFromGUID2(cls, c, 40);
        StringFromGUID2(iid, i, 40);
        stack("CoCreateInstance");
        try {
            logFile << "CLSID=" << std::filesystem::path(c).string()
                    << " IID=" << std::filesystem::path(i).string() << '\n';
        } catch (...) {
        }
    }
    return CoCreateInstance(cls, outer, context, iid, result);
}

BOOL WINAPI writeHook(HANDLE h, LPCVOID bytes, DWORD count, LPDWORD written, LPOVERLAPPED ov) {
    if (GetCurrentThreadId() == traceThread) {
        ++writeCalls;
    }
    if (GetCurrentThreadId() == traceThread && count >= 8 && bytes &&
        !memcmp(bytes, "\x89PNG\r\n\x1a\n", 8)) {
        try {
            wchar_t path[32768]{};
            DWORD n = GetFinalPathNameByHandleW(h, path, 32768, FILE_NAME_NORMALIZED);
            stack("WriteFile PNG payload");
            logFile << "payloadBytes=" << count << " handle=" << h
                    << " offset=" << (ov ? ((uint64_t(ov->OffsetHigh) << 32) | ov->Offset) : 0)
                    << " path="
                    << (n && n < 32768 ? std::filesystem::path(path).string() : "unavailable")
                    << '\n';
            logFile.flush();
        } catch (...) {
        }
    }
    return WriteFile(h, bytes, count, written, ov);
}

DWORD WINAPI typeHook(HANDLE h) {
    if (isMemory(h)) {
        return FILE_TYPE_DISK;
    }
    return GetFileType(h);
}

DWORD WINAPI sizeHook(HANDLE h, LPDWORD high) {
    if (isMemory(h)) {
        if (high) {
            *high = 0;
        }
        SetLastError(0);
        return (DWORD)memoryBytes.size();
    }
    return GetFileSize(h, high);
}

BOOL WINAPI sizeExHook(HANDLE h, PLARGE_INTEGER size) {
    if (isMemory(h)) {
        size->QuadPart = memoryBytes.size();
        return TRUE;
    }
    return GetFileSizeEx(h, size);
}

BOOL WINAPI seekExHook(HANDLE h, LARGE_INTEGER distance, PLARGE_INTEGER result, DWORD method) {
    if (isMemory(h)) {
        long long base = method == FILE_BEGIN     ? 0
                         : method == FILE_CURRENT ? (long long)memoryPosition
                                                  : (long long)memoryBytes.size();
        auto pos = base + distance.QuadPart;
        if (method > FILE_END || pos < 0) {
            SetLastError(ERROR_INVALID_PARAMETER);
            return FALSE;
        }
        memoryPosition = (size_t)pos;
        if (result) {
            result->QuadPart = pos;
        }
        return TRUE;
    }
    return SetFilePointerEx(h, distance, result, method);
}

DWORD WINAPI seekHook(HANDLE h, LONG low, PLONG high, DWORD method) {
    if (isMemory(h)) {
        LARGE_INTEGER d{}, r{};
        d.QuadPart = high ? ((uint64_t((DWORD)*high) << 32) | (DWORD)low) : low;
        if (!seekExHook(h, d, &r, method)) {
            return INVALID_SET_FILE_POINTER;
        }
        if (high) {
            *high = r.HighPart;
        }
        SetLastError(0);
        return r.LowPart;
    }
    return SetFilePointer(h, low, high, method);
}

BOOL WINAPI infoHook(HANDLE h, LPBY_HANDLE_FILE_INFORMATION info) {
    if (isMemory(h)) {
        *info = {};
        info->dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
        info->nFileSizeLow = (DWORD)memoryBytes.size();
        info->nNumberOfLinks = 1;
        info->nFileIndexLow = 0x4242;
        return TRUE;
    }
    return GetFileInformationByHandle(h, info);
}

BOOL WINAPI timeHook(HANDLE h, LPFILETIME a, LPFILETIME b, LPFILETIME c) {
    if (isMemory(h)) {
        FILETIME t{};
        GetSystemTimeAsFileTime(&t);
        if (a) {
            *a = t;
        }
        if (b) {
            *b = t;
        }
        if (c) {
            *c = t;
        }
        return TRUE;
    }
    return GetFileTime(h, a, b, c);
}

BOOL WINAPI overlappedHook(HANDLE h, LPOVERLAPPED ov, LPDWORD bytes, BOOL wait) {
    if (isMemory(h)) {
        *bytes = (DWORD)ov->InternalHigh;
        return TRUE;
    }
    return GetOverlappedResult(h, ov, bytes, wait);
}

struct Patch {
    uintptr_t* slot;
    uintptr_t original;
    uintptr_t replacement;
};

std::vector<Patch> patches;

void change(uintptr_t* slot, uintptr_t expected, uintptr_t replacement) {
    if (*slot != expected) {
        throw bb::Error(E_FAIL, "IAT invariant failed");
    }
    DWORD old = 0;
    if (!VirtualProtect(slot, sizeof(*slot), PAGE_READWRITE, &old)) {
        throw bb::Error(HRESULT_FROM_WIN32(GetLastError()), "IAT protection");
    }
    InterlockedExchangePointer((void* volatile*)slot, (void*)replacement);
    DWORD ignored;
    VirtualProtect(slot, sizeof(*slot), old, &ignored);
}

void restore() {
    for (auto i = patches.rbegin(); i != patches.rend(); ++i) {
        if (*i->slot == i->replacement) {
            change(i->slot, i->replacement, i->original);
        }
    }
    patches.clear();
}

void install() {
    HMODULE mods[1024];
    DWORD bytes;
    bb::check(EnumProcessModules(GetCurrentProcess(), mods, sizeof(mods), &bytes) ? S_OK : E_FAIL,
              "Modules");
    for (unsigned i = 0; i < bytes / sizeof(HMODULE) && i < 1024; i++) {
        wchar_t path[32768]{};
        GetModuleFileNameW(mods[i], path, 32768);
        std::wstring lower = path;
        for (auto& c : lower) {
            c = towlower(c);
        }
        if (lower.find(L"microsoft office") == std::wstring::npos &&
            lower.find(L"microsoft shared\\office") == std::wstring::npos &&
            lower.find(L"windowscodecs.dll") == std::wstring::npos &&
            lower.find(L"gdiplus.dll") == std::wstring::npos) {
            continue;
        }
        auto base = (BYTE*)mods[i];
        auto dos = (IMAGE_DOS_HEADER*)base;
        if (dos->e_magic != IMAGE_DOS_SIGNATURE) {
            continue;
        }
        auto nt = (IMAGE_NT_HEADERS64*)(base + dos->e_lfanew);
        if (nt->Signature != IMAGE_NT_SIGNATURE ||
            nt->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64) {
            continue;
        }
        auto size = nt->OptionalHeader.SizeOfImage;
        auto dir = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
        if (!dir.VirtualAddress || dir.VirtualAddress + dir.Size > size) {
            continue;
        }
        auto desc = (IMAGE_IMPORT_DESCRIPTOR*)(base + dir.VirtualAddress);
        for (; (BYTE*)(desc + 1) <= base + dir.VirtualAddress + dir.Size && desc->Name; desc++) {
            if (!desc->FirstThunk || desc->FirstThunk >= size) {
                continue;
            }
            auto slot = (uintptr_t*)(base + desc->FirstThunk);
            for (; (BYTE*)(slot + 1) <= base + size && *slot; slot++) {
                uintptr_t replacement = 0;
                if (*slot == (uintptr_t)realOpen) {
                    replacement = (uintptr_t)openHook;
                }
                if (*slot == (uintptr_t)realRead) {
                    replacement = (uintptr_t)readHook;
                }
                if (*slot == (uintptr_t)realClose) {
                    replacement = (uintptr_t)closeHook;
                }
                if (*slot == (uintptr_t)WriteFile) {
                    replacement = (uintptr_t)writeHook;
                }
                if (!memoryMode && *slot == (uintptr_t)CreateStreamOnHGlobal) {
                    replacement = (uintptr_t)streamHook;
                }
                if (!memoryMode && *slot == (uintptr_t)CoCreateInstance) {
                    replacement = (uintptr_t)coCreateHook;
                }
                if (memoryMode) {
                    if (*slot == (uintptr_t)GetFileType) {
                        replacement = (uintptr_t)typeHook;
                    }
                    if (*slot == (uintptr_t)GetFileSize) {
                        replacement = (uintptr_t)sizeHook;
                    }
                    if (*slot == (uintptr_t)GetFileSizeEx) {
                        replacement = (uintptr_t)sizeExHook;
                    }
                    if (*slot == (uintptr_t)SetFilePointer) {
                        replacement = (uintptr_t)seekHook;
                    }
                    if (*slot == (uintptr_t)SetFilePointerEx) {
                        replacement = (uintptr_t)seekExHook;
                    }
                    if (*slot == (uintptr_t)GetFileInformationByHandle) {
                        replacement = (uintptr_t)infoHook;
                    }
                    if (*slot == (uintptr_t)GetFileTime) {
                        replacement = (uintptr_t)timeHook;
                    }
                    if (*slot == (uintptr_t)GetOverlappedResult) {
                        replacement = (uintptr_t)overlappedHook;
                    }
                }
                if (replacement) {
                    logFile << "IAT " << std::filesystem::path(path).filename().string()
                            << " slotRVA=0x" << std::hex << ((BYTE*)slot - base) << std::dec
                            << '\n';
                    patches.push_back({slot, *slot, replacement});
                    change(slot, *slot, replacement);
                }
            }
        }
    }
    logFile << "Patched slots=" << patches.size() << '\n';
    logFile.flush();
}
} // namespace

HRESULT memoryFillExperiment(IDispatch* shape, SAFEARRAY* bytes, const std::wstring& root) {
    HRESULT result = S_OK;
    try {
        if (!bytes || SafeArrayGetDim(bytes) != 1) {
            throw bb::Error(E_INVALIDARG, "Expected one-dimensional Byte array");
        }
        VARTYPE vt;
        bb::check(SafeArrayGetVartype(bytes, &vt), "Array type");
        if (vt != VT_UI1) {
            throw bb::Error(E_INVALIDARG, "Expected VT_UI1");
        }
        LONG lo, hi;
        SafeArrayGetLBound(bytes, 1, &lo);
        SafeArrayGetUBound(bytes, 1, &hi);
        if (hi < lo || uint64_t(hi) - lo + 1 > 64 * 1024 * 1024) {
            throw bb::Error(E_INVALIDARG, "Image size must be 1..64 MiB");
        }
        memoryBytes.resize((size_t)(hi - lo + 1));
        void* data;
        bb::check(SafeArrayAccessData(bytes, &data), "Array data");
        memcpy(memoryBytes.data(), data, memoryBytes.size());
        SafeArrayUnaccessData(bytes);
        bb::check(bb::validateImage(memoryBytes.data(), memoryBytes.size()),
                  "Invalid or unsupported PNG/JPEG");
        auto file =
            std::filesystem::path(root) / L"artifacts/BB_TRACE_UNIQUE_TEXTURE_01_MEMORY_ONLY.png";
        file.make_preferred();
        if (std::filesystem::exists(file)) {
            throw bb::Error(E_INVALIDARG,
                            "Memory experiment requires nonexistent sentinel filename");
        }
        auto office = GetModuleHandleW(L"oart.dll");
        if (!office) {
            throw bb::Error(E_NOTIMPL, "No OART module");
        }
        wchar_t modulePath[32768]{};
        GetModuleFileNameW(office, modulePath, 32768);
        DWORD ignored = 0;
        auto versionSize = GetFileVersionInfoSizeW(modulePath, &ignored);
        std::vector<BYTE> version(versionSize);
        if (!versionSize || !GetFileVersionInfoW(modulePath, 0, versionSize, version.data())) {
            throw bb::Error(E_NOTIMPL, "Cannot verify Office version");
        }
        VS_FIXEDFILEINFO* info = nullptr;
        UINT length = 0;
        if (!VerQueryValueW(version.data(), L"\\", (void**)&info, &length) ||
            length < sizeof(*info) || info->dwFileVersionMS != 0x100000 ||
            info->dwFileVersionLS != ((14334u << 16) | 20848u)) {
            throw bb::Error(E_NOTIMPL,
                            "Memory adapter profile supports only Office 16.0.14334.20848");
        }
        const BYTE signature[] = {
            0x40, 0x55, 0x53, 0x56, 0x57, 0x41, 0x56, 0x41, 0x57, 0x48, 0x8d, 0x6c, 0x24, 0xd9};
        auto base = (BYTE*)office;
        auto dos = (IMAGE_DOS_HEADER*)base;
        auto pe = (IMAGE_NT_HEADERS64*)(base + dos->e_lfanew);
        if (pe->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 ||
            pe->OptionalHeader.SizeOfImage < 0x321750 + sizeof(signature) ||
            memcmp(base + 0x321750, signature, sizeof signature)) {
            throw bb::Error(E_NOTIMPL, "Memory adapter OART profile signature mismatch");
        }
        auto fill = bb::get(shape, L"Fill");
        traceThread = GetCurrentThreadId();
        memoryName = file.wstring();
        memoryMode = true;
        logFile.open(std::filesystem::path(root) / L"artifacts/memory_trace.txt");
        install();
        try {
            bb::call(fill.obj(), L"UserPicture", {bb::Value(file.c_str())});
        } catch (...) {
            restore();
            throw;
        }
        restore();
        logFile << "Memory fill returned successfully; no source file exists="
                << !std::filesystem::exists(file) << '\n';
    } catch (const bb::Error& e) {
        result = e.hr;
        logFile << "FAILED " << e.what() << " HRESULT=" << std::hex << (unsigned)e.hr << '\n';
    } catch (...) {
        result = E_UNEXPECTED;
    }
    restore();
    if (tracked != INVALID_HANDLE_VALUE) {
        realClose(tracked);
        tracked = INVALID_HANDLE_VALUE;
    }
    memoryMode = false;
    memoryBytes.clear();
    logFile.close();
    return result;
}

HRESULT traceUserPicture(IDispatch* fill, const std::wstring& root) {
    try {
        logFile.open(std::filesystem::path(root) / L"artifacts/userpicture_trace.txt");
        traceThread = GetCurrentThreadId();
        install();
        try {
            bb::call(fill,
                     L"UserPicture",
                     {bb::Value((std::filesystem::path(root) /
                                 L"artifacts/textures/BB_TRACE_UNIQUE_TEXTURE_01.png")
                                    .c_str())});
        } catch (...) {
            restore();
            throw;
        }
        restore();
        logFile << "All IAT patches restored\n";
        logFile.close();
        return S_OK;
    } catch (const bb::Error& e) {
        restore();
        logFile << e.what() << '\n';
        logFile.close();
        return e.hr;
    } catch (...) {
        restore();
        logFile.close();
        return E_UNEXPECTED;
    }
}

HRESULT traceCachedApply(IDispatch* donor, IDispatch* target, const std::wstring& root) {
    try {
        logFile.open(std::filesystem::path(root) / L"artifacts/cached_apply_trace.txt");
        traceThread = GetCurrentThreadId();
        cacheHandles.clear();
        openCalls = readCalls = writeCalls = 0;
        install();
        try {
            for (int i = 0; i < 100; i++) {
                bb::call(donor, L"PickUp");
                bb::call(target, L"Apply");
            }
        } catch (...) {
            restore();
            throw;
        }
        restore();
        logFile << "100 cached applications completed; hooks restored\n"
                << "monitored openCalls=" << openCalls << " readCalls=" << readCalls
                << " writeCalls=" << writeCalls << '\n';
        logFile.close();
        return S_OK;
    } catch (const bb::Error& e) {
        restore();
        logFile << e.what() << '\n';
        logFile.close();
        return e.hr;
    } catch (...) {
        restore();
        logFile.close();
        return E_UNEXPECTED;
    }
}
