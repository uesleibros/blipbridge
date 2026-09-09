/*
 * BlipBridge - fast reusable image textures for PowerPoint Shapes.
 *
 * This is the stable public interface. It is plain C: no C++ types, no STL
 * types, no exceptions and no ownership of anything the caller allocated ever
 * crosses this boundary. It is what the VBA wrapper binds to, and it is the only
 * interface promised to remain source-compatible.
 *
 * Threading
 * ---------
 * Single-threaded apartment only. BB_Init records the calling thread and every
 * later call must come from that same thread; anything else returns
 * BB_E_WRONG_THREAD rather than corrupting Office state. This is not a
 * conservative choice: the Office objects involved use non-atomic reference
 * counts.
 *
 * Ownership
 * ---------
 *  - Texture handles are owned by the library. Release each with
 *    BB_ReleaseTexture, or release them all with BB_ClearTextures. BB_Shutdown
 *    releases anything still outstanding.
 *  - Handles are never recycled, so a stale handle is reported rather than
 *    silently resolving to a different texture.
 *  - Shape pointers are **borrowed for the duration of the call only**. The
 *    library never stores one. The caller must keep the Shape alive across the
 *    call, which is automatic when passing VBA's ObjPtr(shape) as an argument.
 *  - Byte buffers passed to BB_LoadTexture are copied before the call returns.
 *
 * Errors
 * ------
 * Every entry point returns a BB_Result. On failure, BB_GetLastError gives a
 * human-readable description of the most recent failure on the calling thread.
 *
 * Platform
 * --------
 * The accelerated backend is Windows x64 PowerPoint on specific Office builds.
 * On any other host or build the library loads and answers questions normally,
 * and the texture entry points return BB_E_UNSUPPORTED_BUILD or
 * BB_E_UNSUPPORTED_HOST. It never guesses at an unknown Office layout.
 */

#ifndef BLIPBRIDGE_H
#define BLIPBRIDGE_H

#include <stdint.h>

#if defined(_WIN32)
#  if defined(BLIPBRIDGE_BUILD)
#    define BB_API __declspec(dllexport)
#  else
#    define BB_API __declspec(dllimport)
#  endif
#else
#  define BB_API __attribute__((visibility("default")))
#endif

/*
 * Calling convention.
 *
 * On Windows x64 there is exactly one, so this is empty.
 *
 * On Windows x86 there are several, and the choice is forced: VBA's `Declare`
 * can only call **stdcall**, with no syntax to request anything else. A cdecl
 * export would appear to bind and then leave the stack unbalanced on every
 * call - the worst kind of failure, because it is silent until it is not. So
 * x86 is stdcall.
 *
 * Exported names stay **undecorated** on both, which is what lets one VBA
 * `Declare` name work on either architecture. The 32-bit link step passes
 * `-Wl,--kill-at` to strip the `@N` suffix stdcall would otherwise add; without
 * it the export would be `_BB_Init@0` and no `Declare` would find it.
 */
#if defined(_WIN32) && !defined(_WIN64)
#  define BB_CALL __stdcall
#else
#  define BB_CALL
#endif

#ifdef __cplusplus
extern "C" {
#endif

/**
 * ABI version of this header.
 *
 * Bumped whenever the exported surface changes in a way an older caller could
 * not survive: a signature change, a removed entry point, or a changed meaning.
 * Adding a new export does not bump it, because an older caller simply will not
 * call the new one.
 *
 * A wrapper should compare this against BB_GetAbiVersion() at startup and refuse
 * to continue on a mismatch, so an old .bas paired with a new DLL fails with a
 * clear message instead of calling something whose shape it has wrong.
 */
/*
 * Library version - the release number, and the single place it is written.
 *
 * The ABI implementation and the build system both read these, so the three
 * cannot drift apart the way they had. This is *not* the ABI version: see
 * BB_ABI_VERSION just below, which moves only when the exported contract
 * changes, and docs/c_abi.md for why the two are separate.
 */
#define BB_VERSION_MAJOR 0
#define BB_VERSION_MINOR 4
#define BB_VERSION_PATCH 0

#define BB_ABI_VERSION 2u

/**
 * Opaque texture handle.
 *
 * A 64-bit token, always, on every platform. It is **not** a pointer and must
 * not be dereferenced, cast to one, or given meaning by arithmetic; the only
 * guarantees are that zero is never valid and that values are never recycled
 * within a process.
 *
 * Lifecycle:
 *
 *     BB_LoadTexture / BB_LoadTexturePixels
 *         -> use zero or more times with BB_ApplyTexture
 *         -> BB_ReleaseTexture      (or BB_ClearTextures, or BB_Shutdown)
 *
 * After release the handle is stale for the rest of the process. Using one
 * returns BB_E_INVALID_HANDLE; it never silently resolves to another texture.
 */
typedef uint64_t BB_Handle;

/** Result of every entry point. Zero is success; negative values are failures. */
typedef int32_t BB_Result;

#define BB_OK                     0   /* succeeded                              */
#define BB_E_INVALID_ARG         -1   /* a null or out-of-range argument        */
#define BB_E_NOT_INITIALIZED     -2   /* BB_Init has not run on this thread     */
#define BB_E_WRONG_THREAD        -3   /* called from a thread other than Init's */
#define BB_E_UNSUPPORTED_HOST    -4   /* not running inside PowerPoint          */
#define BB_E_UNSUPPORTED_BUILD   -5   /* Office build not validated             */
#define BB_E_INVALID_HANDLE      -6   /* unknown or already released handle     */
#define BB_E_INVALID_SHAPE       -7   /* not a Shape this backend can fill      */
#define BB_E_DECODE_FAILED       -8   /* the image bytes could not be decoded   */
#define BB_E_APPLY_FAILED        -9   /* the fill could not be applied          */
#define BB_E_OUT_OF_MEMORY      -10   /* allocation failed                      */
#define BB_E_UNSUPPORTED_SHAPE  -11   /* Shape class has no picture-fill path    */
#define BB_E_FILE_NOT_FOUND     -12   /* the image path could not be read        */
#define BB_E_FALLBACK_FAILED    -13   /* Fill.UserPicture itself refused         */
#define BB_E_INTERNAL           -99   /* unexpected internal failure            */

/** Capability bits returned by BB_GetCapabilities. */
#define BB_CAP_NATIVE_BACKEND   0x0001u /* the accelerated backend is usable    */
#define BB_CAP_MEMORY_IMAGE     0x0002u /* bytes reach a fill with no temp file */
#define BB_CAP_CACHED_TEXTURE   0x0004u /* one decode serves many applies       */
#define BB_CAP_BATCH_APPLY      0x0008u /* BB_ApplyTextureBatch is implemented  */
#define BB_CAP_PICKUP_FALLBACK  0x0010u /* the donor COM fallback exists        */
#define BB_CAP_RAW_PIXELS       0x0020u /* BB_LoadTexturePixels is implemented  */
#define BB_CAP_APPLY_PICTURE    0x0040u /* BB_ApplyPicture and its caches exist */

/**
 * Prepares the library on the calling thread and probes the host.
 *
 * Returns BB_OK when the accelerated backend is available. Returns
 * BB_E_UNSUPPORTED_HOST or BB_E_UNSUPPORTED_BUILD when it is not - the library
 * stays usable for BB_GetCapabilities, BB_GetVersion and BB_GetLastError, but
 * the texture entry points will refuse.
 *
 * Calling BB_Init again from the same thread is harmless and re-probes.
 */
BB_API BB_Result BB_CALL BB_Init(void);

/**
 * Releases every texture and forgets the owning thread.
 * Safe to call without a matching BB_Init, and safe to call twice.
 */
BB_API BB_Result BB_CALL BB_Shutdown(void);

/**
 * Decodes @p bytes into a reusable texture.
 *
 * @param bytes  first byte of the encoded image (PNG, JPEG, ...). Copied.
 * @param length number of bytes.
 * @param out    receives the handle; set to 0 on failure.
 */
BB_API BB_Result BB_CALL BB_LoadTexture(const uint8_t* bytes, uint32_t length,
                                        BB_Handle* out);

/**
 * Builds a texture from raw pixels, skipping image decoding entirely.
 *
 * Pixels are 32-bit BGRA - blue, green, red, alpha - which is the ordinary
 * Windows in-memory layout. Only that layout is accepted: the underlying
 * creator takes a surface-format argument, but probing every value from 0 to 24
 * produced an identical BGRA render, so no other layout can be honestly
 * advertised. Convert other layouts before calling.
 *
 * @param pixels first byte of the top-left pixel. Copied by the call.
 * @param width  in pixels.
 * @param height in pixels.
 * @param stride bytes per row; must be at least width*4, and may be larger.
 * @param out    receives the handle; set to 0 on failure.
 *
 * The handle behaves exactly like one from BB_LoadTexture in every other way.
 */
BB_API BB_Result BB_CALL BB_LoadTexturePixels(const uint8_t* pixels, uint32_t width,
                                              uint32_t height, int32_t stride,
                                              BB_Handle* out);

/**
 * Fills one Shape with a texture.
 *
 * @param shape   the Shape's IDispatch pointer - VBA's ObjPtr(shape). Borrowed
 *                for the duration of the call and never stored.
 * @param texture a handle from BB_LoadTexture.
 */
BB_API BB_Result BB_CALL BB_ApplyTexture(void* shape, BB_Handle texture);

/**
 * Fills many Shapes in one call, to avoid a language-boundary crossing per
 * Shape.
 *
 * @param shapes   array of @p count Shape IDispatch pointers.
 * @param textures array of @p count handles, one per Shape. Repeat a handle to
 *                 apply the same texture to several Shapes.
 * @param count    number of entries in both arrays.
 * @param applied  optional; receives how many Shapes were filled before any
 *                 failure. Pass NULL if not wanted.
 *
 * Stops at the first failure and reports it; Shapes already filled stay filled.
 */
BB_API BB_Result BB_CALL BB_ApplyTextureBatch(void* const* shapes,
                                              const BB_Handle* textures,
                                              uint32_t count, uint32_t* applied);

/** Releases one texture. A handle is never valid again afterwards. */
BB_API BB_Result BB_CALL BB_ReleaseTexture(BB_Handle texture);

/** Releases every texture this library holds. */
BB_API BB_Result BB_CALL BB_ClearTextures(void);

/** Number of textures currently held. */
BB_API uint32_t BB_CALL BB_GetTextureCount(void);

/**
 * Capability bits for the *current process*, not a compile-time promise: the
 * accelerated backend is validated against specific Office builds and reports
 * nothing on others.
 */
BB_API uint32_t BB_CALL BB_GetCapabilities(void);

/**
 * Copies a human-readable description of the most recent failure on this
 * thread into @p buffer as UTF-8.
 *
 * @return the number of bytes the message needs including its terminator, so a
 *         caller can size a buffer. Nothing is written when @p buffer is NULL
 *         or @p capacity is 0; the result is always terminated otherwise.
 */
BB_API uint32_t BB_CALL BB_GetLastError(char* buffer, uint32_t capacity);

/**
 * The one-call picture fill: give it a Shape and a file, and it decides.
 *
 * This is the `UserPicture2` the VBA wrapper exposes, and the entry point most
 * callers should use. It removes the need to know anything about Shape classes:
 *
 *   - a Shape class with a validated native path gets the accelerated apply,
 *     from a texture cached by file path so the file is read and decoded once;
 *   - a class without one, but which ordinary `Fill.UserPicture` accepts, gets
 *     that instead;
 *   - anything else returns BB_E_UNSUPPORTED_SHAPE, naming the class.
 *
 * @p path is a **UTF-16, null-terminated** absolute path. VBA passes `StrPtr(s)`
 * directly, with no conversion; UTF-16 is what Windows file APIs take and what
 * VBA already holds, so nothing can be lost in translation.
 *
 * ## The fallback is for Shape classes, not for bugs
 *
 * Falling back happens only when the Shape's *class* has no native path. A
 * native failure caused by an unvalidated Office build, a deleted Shape, or a
 * corrupt image is returned as itself - BB_E_UNSUPPORTED_BUILD,
 * BB_E_INVALID_SHAPE, BB_E_DECODE_FAILED - and never quietly papered over with a
 * slower path that would hide it.
 *
 * ## Caching, and when it is skipped
 *
 * Two caches, both automatic:
 *
 *   - **path to texture.** The file is read and decoded once. The key includes
 *     the file's size and last-write time, so editing the file on disk produces
 *     a new texture rather than a stale one; that check costs about two
 *     microseconds against roughly 190 for an apply.
 *   - **Shape to last texture.** Applying the same image to the same Shape twice
 *     in a row does no Office work at all. Before skipping, `Fill.Type` is
 *     checked to still be a picture fill, so a Shape whose fill was replaced
 *     elsewhere is re-applied rather than left wrong.
 *
 * The second cache cannot see every change. If something replaces the fill with
 * a *different picture* outside this call, the skip will not notice, because
 * detecting it would cost more than the apply it saves. Call
 * BB_InvalidateShape after doing that, or BB_ClearPictureCache after anything
 * wholesale. Both caches are dropped by BB_Shutdown.
 *
 * Returns BB_OK on success, or the specific reason it failed.
 */
BB_API BB_Result BB_CALL BB_ApplyPicture(void* shape, const uint16_t* path);

/**
 * Forgets what was last applied to @p shape, so the next BB_ApplyPicture on it
 * does real work.
 *
 * Call it after changing a Shape's fill by any other means. Forgetting an
 * unknown Shape is not an error.
 */
BB_API BB_Result BB_CALL BB_InvalidateShape(void* shape);

/**
 * Empties both caches: every path-keyed texture is released and every Shape's
 * last-applied record is dropped.
 *
 * Reach for this after closing or reloading a presentation, or when a batch of
 * source files has changed.
 */
BB_API BB_Result BB_CALL BB_ClearPictureCache(void);

/**
 * Cache statistics, for confirming the caches are doing what you think.
 *
 * @p textures receives the number of path-keyed textures held, @p shapes the
 * number of Shapes with a remembered fill, @p skipped the running total of
 * applies avoided since BB_Init. Any pointer may be NULL.
 */
BB_API BB_Result BB_CALL BB_GetPictureCacheStats(uint32_t* textures, uint32_t* shapes,
                                                 uint64_t* skipped);

/**
 * ABI version this DLL implements; compare against BB_ABI_VERSION.
 * Unlike BB_GetVersion this is not a release number - it changes only when the
 * exported surface stops being compatible.
 */
BB_API uint32_t BB_CALL BB_GetAbiVersion(void);

/** Packed release version: (major << 16) | (minor << 8) | patch. */
BB_API uint32_t BB_CALL BB_GetVersion(void);

/**
 * Copies the version and build description as UTF-8, for example
 * "0.2.0 (windows-x64)". Same return convention as BB_GetLastError.
 */
BB_API uint32_t BB_CALL BB_GetVersionString(char* buffer, uint32_t capacity);

#ifdef __cplusplus
} /* extern "C" */
#endif

#endif /* BLIPBRIDGE_H */
