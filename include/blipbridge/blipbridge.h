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
 * Calling convention. On Windows x64 there is exactly one, so this is empty and
 * exported names are undecorated - which is what lets VBA's Declare bind to them
 * by plain name. It is spelled out so a future 32-bit or non-Windows target
 * cannot silently change the ABI.
 */
#define BB_CALL

#ifdef __cplusplus
extern "C" {
#endif

/** Opaque texture handle. Zero is never a valid handle. */
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
#define BB_E_INTERNAL           -99   /* unexpected internal failure            */

/** Capability bits returned by BB_GetCapabilities. */
#define BB_CAP_NATIVE_BACKEND   0x0001u /* the accelerated backend is usable    */
#define BB_CAP_MEMORY_IMAGE     0x0002u /* bytes reach a fill with no temp file */
#define BB_CAP_CACHED_TEXTURE   0x0004u /* one decode serves many applies       */
#define BB_CAP_BATCH_APPLY      0x0008u /* BB_ApplyTextureBatch is implemented  */
#define BB_CAP_PICKUP_FALLBACK  0x0010u /* the donor COM fallback exists        */

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

/** Packed version: (major << 16) | (minor << 8) | patch. */
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
