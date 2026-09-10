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
#include <cstring>
#include <iostream>
#include <string>
#include <vector>
#include <windows.h>

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
/*
 * Every pointer below carries BB_CALL, and that is not decoration.
 *
 * On x64 it expands to nothing, because there is one convention. On x86 it is
 * __stdcall, which is what the DLL exports - and a cdecl pointer to a stdcall
 * function leaves the stack unbalanced on every call. The failure does not
 * appear at the call site; it appears somewhere afterwards, as a crash that
 * looks unrelated. CI found exactly that when these were written without it.
 */
struct Api {
    HMODULE module = nullptr;

    BB_Result(BB_CALL* Init)(void) = nullptr;
    BB_Result(BB_CALL* Shutdown)(void) = nullptr;
    BB_Result(BB_CALL* LoadTexture)(const uint8_t*, uint32_t, BB_Handle*) = nullptr;
    BB_Result(BB_CALL* ApplyTexture)(void*, BB_Handle) = nullptr;
    BB_Result(BB_CALL* ApplyTextureBatch)(void* const*,
                                          const BB_Handle*,
                                          uint32_t,
                                          uint32_t*) = nullptr;
    BB_Result(BB_CALL* ReleaseTexture)(BB_Handle) = nullptr;
    BB_Result(BB_CALL* ClearTextures)(void) = nullptr;
    uint32_t(BB_CALL* GetTextureCount)(void) = nullptr;
    uint32_t(BB_CALL* GetCapabilities)(void) = nullptr;
    uint32_t(BB_CALL* GetLastError)(char*, uint32_t) = nullptr;
    uint32_t(BB_CALL* GetVersion)(void) = nullptr;
    uint32_t(BB_CALL* GetAbiVersion)(void) = nullptr;
    BB_Result(BB_CALL* LoadTexturePixels)(const uint8_t*, uint32_t, uint32_t, int32_t, BB_Handle*) =
        nullptr;
    uint32_t(BB_CALL* GetVersionString)(char*, uint32_t) = nullptr;
    BB_Result(BB_CALL* ApplyPicture)(void*, const uint16_t*) = nullptr;
    BB_Result(BB_CALL* InvalidateShape)(void*) = nullptr;
    BB_Result(BB_CALL* ClearPictureCache)(void) = nullptr;
    BB_Result(BB_CALL* GetPictureCacheStats)(uint32_t*, uint32_t*, uint64_t*) = nullptr;

    ~Api() {
        if (module) {
            FreeLibrary(module);
        }
    }
};

template <typename Fn>
void Resolve(Api& api, Fn& slot, const char* name) {
    slot = reinterpret_cast<Fn>(reinterpret_cast<void*>(GetProcAddress(api.module, name)));
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
    Resolve(api, api.GetAbiVersion, "BB_GetAbiVersion");
    Resolve(api, api.LoadTexturePixels, "BB_LoadTexturePixels");
    Resolve(api, api.GetVersionString, "BB_GetVersionString");
    Resolve(api, api.ApplyPicture, "BB_ApplyPicture");
    Resolve(api, api.InvalidateShape, "BB_InvalidateShape");
    Resolve(api, api.ClearPictureCache, "BB_ClearPictureCache");
    Resolve(api, api.GetPictureCacheStats, "BB_GetPictureCacheStats");
    if (g_failures) {
        return 1;
    }

    // --- questions answerable without any host --------------------------------
    Check(api.GetVersion() != 0, "BB_GetVersion returns a non-zero packed version");
    Check(api.GetAbiVersion() == BB_ABI_VERSION,
          "the DLL's ABI version matches the header this test was built against");
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
    {
        const uint16_t path[] = {L'x', 0};
        Check(api.ApplyPicture(nullptr, path) == BB_E_NOT_INITIALIZED,
              "ApplyPicture before Init reports BB_E_NOT_INITIALIZED");
        Check(api.InvalidateShape(nullptr) == BB_E_NOT_INITIALIZED,
              "InvalidateShape before Init reports BB_E_NOT_INITIALIZED");
        Check(api.ClearPictureCache() == BB_E_NOT_INITIALIZED,
              "ClearPictureCache before Init reports BB_E_NOT_INITIALIZED");
        Check(api.GetPictureCacheStats(nullptr, nullptr, nullptr) == BB_E_NOT_INITIALIZED,
              "GetPictureCacheStats before Init reports BB_E_NOT_INITIALIZED");
    }
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
    Check(api.ApplyTexture(nullptr, 1) == BB_E_INVALID_ARG, "ApplyTexture rejects a null Shape");
    Check(api.ApplyTexture(reinterpret_cast<void*>(&byte), 0) == BB_E_INVALID_HANDLE,
          "handle 0 is never valid");
    Check(api.ReleaseTexture(0) == BB_E_INVALID_HANDLE, "ReleaseTexture rejects handle 0");

    /*
     * The 64-bit handle must survive the ABI on a 32-bit build too.
     *
     * BB_Handle is uint64_t on every architecture, and the ABI does not shrink
     * it to suit a 32-bit host. On x86 that means it crosses the boundary as two
     * stack slots, and the VBA wrapper hands it over as two Longs. If the
     * calling convention or the handle width were wrong, a value with bits set
     * in *both* halves would be misread, and the stack would be left unbalanced
     * - so the next call would misbehave rather than this one.
     *
     * Hence the shape of this check: pass such a value, and then confirm the
     * library still answers correctly afterwards.
     */
    {
        Check(sizeof(BB_Handle) == 8, "BB_Handle is 64 bits on this architecture");

        const BB_Handle wide = 0x1234abcd5678ef01ull;
        /*
         * Which refusal comes back depends on whether this build has a backend:
         * one that tracks handles says BB_E_INVALID_HANDLE, one that has no
         * backend at all says BB_E_UNSUPPORTED_HOST first. Both are correct, and
         * pinning either would make this test assert the architecture rather
         * than the thing it is here for - that a 64-bit value crosses the ABI
         * intact. What must not happen is success.
         */
        const BB_Result refused = api.ReleaseTexture(wide);
        Check(refused != BB_OK, "a handle with both halves set is refused, not misread as valid");
        // A null Shape, deliberately: it is refused by argument validation before
        // anything dereferences it. A non-null fake would be dereferenced once the
        // handle check no longer short-circuits, and that is a crash, not a test.
        Check(api.ApplyTexture(nullptr, wide) == BB_E_INVALID_ARG,
              "the same handle reaches ApplyTexture's argument checks intact");
        Check(!LastError(api).empty(), "and the refusal came with a readable reason");

        // The real assertion: the stack survived. A convention mismatch shows up
        // here rather than above.
        Check(api.GetVersion() != 0, "the library still answers after a 64-bit handle call");
        Check(api.GetTextureCount() == 0, "and still reports a consistent texture count");
        Check(api.GetAbiVersion() == BB_ABI_VERSION, "and still reports its ABI version");
    }
    {
        const uint16_t path[] = {L'x', 0};
        const uint16_t empty[] = {0};
        Check(api.ApplyPicture(nullptr, path) == BB_E_INVALID_ARG,
              "ApplyPicture rejects a null Shape");
        Check(api.ApplyPicture(reinterpret_cast<void*>(&byte), nullptr) == BB_E_INVALID_ARG,
              "ApplyPicture rejects a null path");
        Check(api.ApplyPicture(reinterpret_cast<void*>(&byte), empty) == BB_E_INVALID_ARG,
              "ApplyPicture rejects an empty path");
        Check(api.InvalidateShape(nullptr) == BB_E_INVALID_ARG,
              "InvalidateShape rejects a null Shape");

        // Every out pointer is optional, so asking for nothing must still work.
        Check(api.GetPictureCacheStats(nullptr, nullptr, nullptr) == BB_OK,
              "GetPictureCacheStats accepts all-null outputs");
        uint32_t textures = 99;
        uint32_t shapes = 99;
        uint64_t skipped = 99;
        Check(api.GetPictureCacheStats(&textures, &shapes, &skipped) == BB_OK,
              "GetPictureCacheStats succeeds");
        Check(textures == 0 && shapes == 0 && skipped == 0, "an untouched cache reports zeroes");
    }

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

    // --- raw pixel argument validation ---------------------------------------
    const uint8_t bgra[16] = {}; // 2x2 BGRA
    Check(api.LoadTexturePixels(nullptr, 2, 2, 8, &handle) == BB_E_INVALID_ARG,
          "pixel load rejects a null buffer");
    Check(api.LoadTexturePixels(bgra, 0, 2, 8, &handle) == BB_E_INVALID_ARG,
          "pixel load rejects zero width");
    Check(api.LoadTexturePixels(bgra, 2, 0, 8, &handle) == BB_E_INVALID_ARG,
          "pixel load rejects zero height");
    Check(api.LoadTexturePixels(bgra, 2, 2, 8, nullptr) == BB_E_INVALID_ARG,
          "pixel load rejects a null output pointer");
    if (!inPowerPoint) {
        Check(api.LoadTexturePixels(bgra, 2, 2, 4, &handle) != BB_OK,
              "a stride below width*4 is refused");
    }

    // --- lifecycle semantics --------------------------------------------------
    // Init is documented as safe to repeat and as re-probing, and Shutdown as
    // safe without a matching Init and safe to repeat. These are the contracts a
    // wrapper relies on when it calls Initialize from every entry point.
    const BB_Result again = api.Init();
    Check(again == init, "a second Init reports the same thing as the first");
    Check(api.Init() == init, "Init stays idempotent");

    Check(api.Shutdown() == BB_OK, "Shutdown succeeds");
    Check(api.Shutdown() == BB_OK, "Shutdown is idempotent");
    Check(api.LoadTexture(&byte, 1, &handle) == BB_E_NOT_INITIALIZED,
          "after Shutdown the library is uninitialised again");
    Check(api.Init() == init, "Init works again after Shutdown");
    Check(api.Shutdown() == BB_OK, "and Shutdown still succeeds after that");

    if (g_failures == 0) {
        std::cout << "ABI contract tests passed\n";
        return 0;
    }
    std::cerr << g_failures << " ABI contract checks failed\n";
    return 1;
}
