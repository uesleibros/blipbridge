/**
 * @file abi_contract.cpp
 * Contract tests for the public C ABI, with no COM Automation involved.
 *
 * These run outside PowerPoint, which is the point: the ABI must behave
 * predictably on a host where the accelerated backend cannot work. That is the
 * common case for anyone who downloads the library, and getting a clear
 * BB_E_UNSUPPORTED_HOST with a readable message is the difference between a
 * usable library and a mysterious one.
 *
 * The DLL is loaded exactly the way the VBA wrapper loads it - LoadLibraryW by
 * full path, then GetProcAddress by undecorated name - so this also checks that
 * the export table is what the wrapper expects.
 */

#include <blipbridge/blipbridge.h>

#include <windows.h>

#include <cstring>
#include <iostream>
#include <string>
#include <vector>

namespace {

int g_failures = 0;

void Check(bool condition, const std::string& what) {
    if (condition) {
        return;
    }
    std::cerr << "FAILED: " << what << '\n';
    ++g_failures;
}

/// Every entry point, resolved by name the way the VBA wrapper resolves them.
struct Api {
    HMODULE module = nullptr;

    BB_Result (*Init)(void) = nullptr;
    BB_Result (*Shutdown)(void) = nullptr;
    BB_Result (*LoadTexture)(const uint8_t*, uint32_t, BB_Handle*) = nullptr;
    BB_Result (*ApplyTexture)(void*, BB_Handle) = nullptr;
    BB_Result (*ApplyTextureBatch)(void* const*, const BB_Handle*, uint32_t,
                                   uint32_t*) = nullptr;
    BB_Result (*ReleaseTexture)(BB_Handle) = nullptr;
    BB_Result (*ClearTextures)(void) = nullptr;
    uint32_t (*GetTextureCount)(void) = nullptr;
    uint32_t (*GetCapabilities)(void) = nullptr;
    uint32_t (*GetLastError)(char*, uint32_t) = nullptr;
    uint32_t (*GetVersion)(void) = nullptr;
    uint32_t (*GetVersionString)(char*, uint32_t) = nullptr;

    ~Api() {
        if (module) {
            FreeLibrary(module);
        }
    }
};

template <typename Fn>
void Resolve(Api& api, Fn& slot, const char* name) {
    slot = reinterpret_cast<Fn>(
        reinterpret_cast<void*>(GetProcAddress(api.module, name)));
    Check(slot != nullptr, std::string("export missing: ") + name);
}

std::string LastError(const Api& api) {
    const uint32_t needed = api.GetLastError(nullptr, 0);
    if (needed <= 1) {
        return {};
    }
    std::vector<char> buffer(needed);
    api.GetLastError(buffer.data(), needed);
    return std::string(buffer.data());
}

} // namespace

int wmain(int argc, wchar_t** argv) {
    if (argc != 2) {
        std::cerr << "Expected the path to BlipBridge.dll\n";
        return 1;
    }

    Api api;
    api.module = LoadLibraryW(argv[1]);
    Check(api.module != nullptr, "LoadLibraryW on the DLL path");
    if (!api.module) {
        return 1;
    }

    Resolve(api, api.Init, "BB_Init");
    Resolve(api, api.Shutdown, "BB_Shutdown");
    Resolve(api, api.LoadTexture, "BB_LoadTexture");
    Resolve(api, api.ApplyTexture, "BB_ApplyTexture");
    Resolve(api, api.ApplyTextureBatch, "BB_ApplyTextureBatch");
    Resolve(api, api.ReleaseTexture, "BB_ReleaseTexture");
    Resolve(api, api.ClearTextures, "BB_ClearTextures");
    Resolve(api, api.GetTextureCount, "BB_GetTextureCount");
    Resolve(api, api.GetCapabilities, "BB_GetCapabilities");
    Resolve(api, api.GetLastError, "BB_GetLastError");
    Resolve(api, api.GetVersion, "BB_GetVersion");
    Resolve(api, api.GetVersionString, "BB_GetVersionString");
    if (g_failures) {
        return 1;
    }

    // --- questions answerable without any host --------------------------------
    Check(api.GetVersion() != 0, "BB_GetVersion returns a non-zero packed version");
    {
        const uint32_t needed = api.GetVersionString(nullptr, 0);
        Check(needed > 1, "BB_GetVersionString reports a required size");
        std::vector<char> buffer(needed);
        const uint32_t written = api.GetVersionString(buffer.data(), needed);
        Check(written == needed, "BB_GetVersionString agrees with its own size");
        Check(std::strlen(buffer.data()) == needed - 1,
              "BB_GetVersionString writes a terminated string");

        // Truncation must still terminate rather than overrun.
        char small[4] = {'x', 'x', 'x', 'x'};
        api.GetVersionString(small, 4);
        Check(small[3] == '\0', "BB_GetVersionString terminates a truncated result");
    }

    // --- calls before BB_Init -------------------------------------------------
    BB_Handle handle = 0xDEADBEEF;
    Check(api.LoadTexture(nullptr, 0, &handle) == BB_E_NOT_INITIALIZED,
          "LoadTexture before Init reports BB_E_NOT_INITIALIZED");
    Check(handle == 0, "LoadTexture zeroes the output handle on failure");
    Check(api.ApplyTexture(nullptr, 1) == BB_E_NOT_INITIALIZED,
          "ApplyTexture before Init reports BB_E_NOT_INITIALIZED");
    Check(api.ReleaseTexture(1) == BB_E_NOT_INITIALIZED,
          "ReleaseTexture before Init reports BB_E_NOT_INITIALIZED");
    Check(!LastError(api).empty(), "a failure leaves a readable message");

    // --- Init on an unsupported host -----------------------------------------
    const BB_Result init = api.Init();
    const bool inPowerPoint = GetModuleHandleW(L"POWERPNT.EXE") != nullptr;
    if (!inPowerPoint) {
        Check(init == BB_E_UNSUPPORTED_HOST,
              "Init outside PowerPoint reports BB_E_UNSUPPORTED_HOST");
        const std::string message = LastError(api);
        Check(message.find("PowerPoint") != std::string::npos,
              "the unsupported-host message names PowerPoint");
        Check((api.GetCapabilities() & BB_CAP_NATIVE_BACKEND) == 0,
              "no native capability is advertised off-host");

        // Init still claims the thread, so later calls give the real reason
        // rather than "not initialised".
        const uint8_t pixels[] = {0x89, 'P', 'N', 'G'};
        Check(api.LoadTexture(pixels, sizeof(pixels), &handle) != BB_E_NOT_INITIALIZED,
              "after Init the failure reason is specific, not 'not initialised'");
        Check(handle == 0, "no handle is produced on an unsupported host");
    }

    // --- argument validation, independent of host ----------------------------
    Check(api.LoadTexture(nullptr, 4, &handle) == BB_E_INVALID_ARG,
          "LoadTexture rejects a null buffer");
    uint8_t byte = 0;
    Check(api.LoadTexture(&byte, 0, &handle) == BB_E_INVALID_ARG,
          "LoadTexture rejects a zero length");
    Check(api.LoadTexture(&byte, 1, nullptr) == BB_E_INVALID_ARG,
          "LoadTexture rejects a null output pointer");
    Check(api.ApplyTexture(nullptr, 1) == BB_E_INVALID_ARG,
          "ApplyTexture rejects a null Shape");
    Check(api.ApplyTexture(reinterpret_cast<void*>(&byte), 0) == BB_E_INVALID_HANDLE,
          "handle 0 is never valid");
    Check(api.ReleaseTexture(0) == BB_E_INVALID_HANDLE,
          "ReleaseTexture rejects handle 0");

    // --- batch argument validation -------------------------------------------
    uint32_t applied = 99;
    Check(api.ApplyTextureBatch(nullptr, nullptr, 0, &applied) == BB_E_INVALID_ARG,
          "batch rejects null arrays");
    Check(applied == 0, "batch zeroes its applied counter up front");
    void* oneShape[1] = {nullptr};
    const BB_Handle oneTexture[1] = {1};
    Check(api.ApplyTextureBatch(oneShape, oneTexture, 0, &applied) == BB_OK,
          "an empty batch succeeds trivially");
    Check(api.ApplyTextureBatch(oneShape, oneTexture, 1, &applied) == BB_E_INVALID_ARG,
          "batch rejects a null Shape entry");
    void* fakeShape[1] = {reinterpret_cast<void*>(&byte)};
    const BB_Handle zeroTexture[1] = {0};
    Check(api.ApplyTextureBatch(fakeShape, zeroTexture, 1, &applied) == BB_E_INVALID_ARG,
          "batch rejects handle 0 in an entry");

    // --- counting and clearing are always answerable -------------------------
    Check(api.GetTextureCount() == 0, "no textures are held on an unsupported host");
    api.ClearTextures();
    Check(api.GetTextureCount() == 0, "ClearTextures leaves nothing behind");

    // --- shutdown ------------------------------------------------------------
    Check(api.Shutdown() == BB_OK, "Shutdown succeeds");
    Check(api.Shutdown() == BB_OK, "Shutdown is idempotent");
    Check(api.LoadTexture(&byte, 1, &handle) == BB_E_NOT_INITIALIZED,
          "after Shutdown the library is uninitialised again");

    if (g_failures == 0) {
        std::cout << "ABI contract tests passed\n";
        return 0;
    }
    std::cerr << g_failures << " ABI contract checks failed\n";
    return 1;
}
