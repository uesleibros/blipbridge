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
#if defined(BLIPBRIDGE_BUILD)
#define BB_API __declspec(dllexport)
#else
#define BB_API __declspec(dllimport)
#endif
#else
#define BB_API __attribute__((visibility("default")))
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
#define BB_CALL __stdcall
#else
#define BB_CALL
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
#define BB_VERSION_MINOR 7
#define BB_VERSION_PATCH 1

#define BB_ABI_VERSION 5u

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

#define BB_OK 0                    /* succeeded                              */
#define BB_E_INVALID_ARG -1        /* a null or out-of-range argument        */
#define BB_E_NOT_INITIALIZED -2    /* BB_Init has not run on this thread     */
#define BB_E_WRONG_THREAD -3       /* called from a thread other than Init's */
#define BB_E_UNSUPPORTED_HOST -4   /* not running inside PowerPoint          */
#define BB_E_UNSUPPORTED_BUILD -5  /* Office build not validated             */
#define BB_E_INVALID_HANDLE -6     /* unknown or already released handle     */
#define BB_E_INVALID_SHAPE -7      /* not a Shape this backend can fill      */
#define BB_E_DECODE_FAILED -8      /* the image bytes could not be decoded   */
#define BB_E_APPLY_FAILED -9       /* the fill could not be applied          */
#define BB_E_OUT_OF_MEMORY -10     /* allocation failed                      */
#define BB_E_UNSUPPORTED_SHAPE -11 /* Shape class has no picture-fill path    */
#define BB_E_FILE_NOT_FOUND -12    /* the image path could not be read        */
#define BB_E_FALLBACK_FAILED -13   /* Fill.UserPicture itself refused         */
#define BB_E_INTERNAL -99          /* unexpected internal failure            */

/** Capability bits returned by BB_GetCapabilities. */
#define BB_CAP_NATIVE_BACKEND 0x0001u  /* the accelerated backend is usable    */
#define BB_CAP_MEMORY_IMAGE 0x0002u    /* bytes reach a fill with no temp file */
#define BB_CAP_CACHED_TEXTURE 0x0004u  /* one decode serves many applies       */
#define BB_CAP_BATCH_APPLY 0x0008u     /* BB_ApplyTextureBatch is implemented  */
#define BB_CAP_PICKUP_FALLBACK 0x0010u /* the donor COM fallback exists        */
#define BB_CAP_RAW_PIXELS 0x0020u      /* BB_LoadTexturePixels is implemented  */
#define BB_CAP_APPLY_PICTURE 0x0040u   /* BB_ApplyPicture and its caches exist */
#define BB_CAP_SCALED_PIXELS 0x0080u   /* BB_LoadTexturePixelsScaled is present */
#define BB_CAP_RANGE_APPLY 0x0100u     /* BB_ApplyTextureRange fills a ShapeRange */
#define BB_CAP_IMAGE_PIPELINE 0x0200u  /* BB_Image, crop/transform/scale, quad warp */

/*
 * Resampling filters for BB_LoadTexturePixelsScaled.
 *
 * Every value here is implemented and tested. A filter is not given a name until
 * it works, so there are no placeholders to discover at run time.
 *
 * These control the image BlipBridge hands to Office. They say nothing about how
 * Office then draws it - Shape scaling, slideshow scaling, zoom and DPI are all
 * downstream and unaffected.
 */
#define BB_SCALE_NEAREST 0u  /* exact point sampling; bytes preserved */
#define BB_SCALE_BILINEAR 1u /* 2x2 interpolation                     */
#define BB_SCALE_BICUBIC 2u  /* Catmull-Rom cubic over a 4x4 window   */

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
BB_API BB_Result BB_CALL BB_LoadTexture(const uint8_t* bytes, uint32_t length, BB_Handle* out);

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
BB_API BB_Result BB_CALL BB_LoadTexturePixels(
    const uint8_t* pixels, uint32_t width, uint32_t height, int32_t stride, BB_Handle* out);

/**
 * Fills one Shape with a texture.
 *
 * @param shape   the Shape's IDispatch pointer - VBA's ObjPtr(shape). Borrowed
 *                for the duration of the call and never stored.
 * @param texture a handle from BB_LoadTexture.
 */
BB_API BB_Result BB_CALL BB_ApplyTexture(void* shape, BB_Handle texture);

/**
 * Applies @p texture to @p shape only if that Shape does not already carry it.
 *
 * Identical to BB_ApplyTexture in what it accepts and in what the document looks
 * like afterwards. The difference is that a Shape already carrying the image is
 * left completely untouched: no document edit, no undo entry, no invalidation,
 * no modified flag. @p skipped, when given, receives 1 in that case and 0 when
 * the fill was really applied.
 *
 * The image is compared by internal identity, not by handle, so two handles for
 * the same picture compare equal and a released handle cannot alias a new one.
 *
 * Because BlipBridge cannot see a fill replaced behind its back by something
 * else - another add-in, a paste, a theme change - use BB_InvalidateShape after
 * such a change, or BB_ApplyTexture, which never skips.
 */
BB_API BB_Result BB_CALL BB_ApplyTextureIfChanged(void* shape,
                                                  BB_Handle texture,
                                                  int32_t* skipped);

/**
 * Applies @p texture to every Shape in one PowerPoint ShapeRange, in one
 * operation.
 *
 * @p shapeRange is a `ShapeRange` - `ObjPtr(Slide.Shapes.Range(...))` from VBA -
 * borrowed for the call and never retained. It is not a `Shape`: a Shape cannot
 * answer `Count` and is refused with a message saying so. @p applied, which may
 * be `NULL`, receives how many member Shapes were filled.
 *
 * This is **not** a loop over BB_ApplyTexture, and it is not BB_ApplyTextureBatch
 * with different arguments. It is one cached texture, one ShapeRange, and one
 * apply through Office's own range receiver - which is what makes it faster
 * rather than merely tidier. Measured on the validated build: 1.876 ms for 32
 * Shapes against 6.650 ms one at a time, 5.549 ms for 100 against 20.935 ms.
 *
 * **All or nothing.** Every member is classified before any internal object is
 * touched. One member without a validated native picture-fill path - a
 * Connector, a Line, a Chart, a Table - refuses the whole call with nothing
 * applied and the private backend never entered. The error names which member.
 *
 * **One slide.** A PowerPoint ShapeRange belongs to one slide's `Shapes`
 * collection and cannot span slides, so neither can this. Filling Shapes on
 * several slides is one call per slide, and the fixed cost is paid per slide.
 *
 * **Undo.** One range apply is one undo entry, and one Undo reverts the whole
 * fill - the same as Office's own `ShapeRange.Fill.UserPicture`. Counted with a
 * marker Shape on the undo stack at 4, 8 and 16 members and on a range of five
 * different Shape classes; filling the same Shapes one at a time leaves one
 * entry each, which is what `BB_ApplyTexture` has always done.
 *
 * Undo and Redo change fills without telling BlipBridge, like any change made
 * outside it, so the per-Shape record used by `BB_ApplyTextureIfChanged` can be
 * stale afterwards. `BB_InvalidateShape` is the remedy, as for any other
 * external change.
 *
 * Groups may be members and are filled. They are deliberately never remembered
 * for the `BB_ApplyTextureIfChanged` skip, because a group's fill and its
 * children's fills change each other.
 */
BB_API BB_Result BB_CALL BB_ApplyTextureRange(void* shapeRange,
                                              BB_Handle texture,
                                              uint32_t* applied);

/* ---------------------------------------------------------------------------
 * Images: the CPU-side resource
 *
 * A BB_Handle texture is an *Office* resource. It holds what Office needs and
 * no pixels, which is why nothing can crop or warp one. Processing therefore
 * gets its own resource: a BB_Image is decoded BGRA32 that BlipBridge owns, on
 * the CPU, with no Office in it.
 *
 * The division is deliberate and neither side leaks into the other:
 *
 *     BB_LoadTextureEx        encoded -> texture      one shot, no repeat work
 *     BB_LoadImage            encoded -> image        decode once
 *     BB_WarpImageQuad        image   -> texture      warp many times
 *
 * A texture never starts retaining pixels behind your back, and an image never
 * touches a Shape. They convert when you ask them to.
 *
 * Image handles come from their own numbering space, so passing a texture handle
 * to an image call is refused rather than resolving to something unrelated.
 * Neither space recycles a released handle.
 * ------------------------------------------------------------------------- */

/** An opaque CPU image handle. Zero is never valid. Release with BB_ReleaseImage. */
typedef uint64_t BB_Image;

/** Orientation changes. Every output pixel is exactly one input pixel. */
#define BB_TRANSFORM_NONE 0u
#define BB_TRANSFORM_FLIP_HORIZONTAL 1u
#define BB_TRANSFORM_FLIP_VERTICAL 2u
#define BB_TRANSFORM_ROTATE_90 3u  /* clockwise */
#define BB_TRANSFORM_ROTATE_180 4u
#define BB_TRANSFORM_ROTATE_270 5u /* clockwise, i.e. 90 anticlockwise */

/**
 * What to do to an image on the way in.
 *
 * The stages run in a fixed order - crop, then transform, then resize - and each
 * switches itself off when unset, so a zeroed request changes nothing. Crop is
 * first because a region is named in the source's own coordinates; transform is
 * before resize so that the target size always describes what comes out.
 *
 * A zero crop width or height means the whole image. A zero target width or
 * height means "whatever the earlier stages produced".
 */
typedef struct BB_ImageRequest {
    uint32_t cropX;
    uint32_t cropY;
    uint32_t cropWidth;
    uint32_t cropHeight;
    uint32_t transform; /* BB_TRANSFORM_* */
    uint32_t targetWidth;
    uint32_t targetHeight;
    uint32_t filter; /* BB_SCALE_* */
} BB_ImageRequest;

/** One quad corner, in the caller's own coordinate space. */
typedef struct BB_PointF {
    float x;
    float y;
} BB_PointF;

/**
 * Decodes @p bytes into a CPU image, applying @p request on the way.
 *
 * @p request may be NULL, which decodes and changes nothing. PNG, JPEG and BMP
 * are tested; Windows decodes more formats and they will probably work, but an
 * untested format is not a supported one.
 *
 * The returned handle is yours. Release it with BB_ReleaseImage.
 */
BB_API BB_Result BB_CALL BB_LoadImage(const uint8_t* bytes,
                                      uint32_t length,
                                      const BB_ImageRequest* request,
                                      BB_Image* out);

/** As BB_LoadImage, reading the file itself. @p path is UTF-16. */
BB_API BB_Result BB_CALL BB_LoadImageFromFile(const uint16_t* path,
                                              const BB_ImageRequest* request,
                                              BB_Image* out);

/** As BB_LoadImage, from raw BGRA32 you already hold. The bytes are copied. */
BB_API BB_Result BB_CALL BB_LoadImagePixels(const uint8_t* pixels,
                                            uint32_t width,
                                            uint32_t height,
                                            int32_t stride,
                                            const BB_ImageRequest* request,
                                            BB_Image* out);

/** The image's dimensions in pixels. Either output pointer may be NULL. */
BB_API BB_Result BB_CALL BB_GetImageSize(BB_Image image, uint32_t* width, uint32_t* height);

/** Releases one image. A released handle stays stale for the life of the process. */
BB_API BB_Result BB_CALL BB_ReleaseImage(BB_Image image);

/** Releases every image. Textures are a different resource and are untouched. */
BB_API BB_Result BB_CALL BB_ClearImages(void);

/** How many images are currently held. */
BB_API uint32_t BB_CALL BB_GetImageCount(void);

/**
 * Creates an Office texture from a CPU image.
 *
 * The image is unchanged and still yours; the texture is a new, separate,
 * caller-owned resource to release with BB_ReleaseTexture. Nothing aliases.
 */
BB_API BB_Result BB_CALL BB_CreateTextureFromImage(BB_Image image, BB_Handle* out);

/**
 * Warps @p image onto four corners and returns it as a texture.
 *
 * This is the primitive. Decode once with BB_LoadImage, then warp as often as
 * the quad moves - the decode is by far the expensive half, and this never
 * repeats it.
 *
 * The points are the destination corners of the source image's own corners, in
 * this order and never reordered behind you:
 *
 *     points[0] = source top-left
 *     points[1] = source top-right
 *     points[2] = source bottom-right
 *     points[3] = source bottom-left
 *
 * Handing them in another order asks for a mirrored or crossed mapping and gets
 * one, which is the only behaviour that lets you mirror something on purpose.
 *
 * The mapping is a true projective transform, so a trapezoid foreshortens the
 * way perspective does rather than the way a shear does. The result is
 * rasterised into the quad's bounding box with everything outside the quad left
 * transparent, and it is sized to that box in the caller's own units.
 *
 * @p filter is BB_SCALE_NEAREST or BB_SCALE_BILINEAR. Bicubic is refused by
 * name: sixteen taps per output pixel with a varying footprint is a cost this
 * would be hiding rather than offering.
 *
 * Coordinates are in whatever unit you work in - only their relative geometry
 * matters - but for the fill to land where you expect, the Shape's own geometry
 * has to be that quad. BlipBridge does not move Shapes, and it does not read
 * Shape.Nodes to find out where they are: you already know the points.
 *
 * Returns a new caller-owned texture. The image is untouched.
 */
BB_API BB_Result BB_CALL BB_WarpImageQuad(BB_Image image,
                                          const BB_PointF* points,
                                          uint32_t filter,
                                          BB_Handle* out);

/**
 * Warps @p image onto four corners and applies it to @p shape in one call.
 *
 * Convenience over BB_WarpImageQuad plus BB_ApplyTexture plus BB_ReleaseTexture,
 * for the common case where the warped result is used once. A caller who applies
 * the same warp to several Shapes should use the primitive and keep the texture.
 *
 * The same Shape-class gate as BB_ApplyTexture runs before anything internal is
 * touched.
 */
BB_API BB_Result BB_CALL BB_ApplyImageQuad(void* shape,
                                           BB_Image image,
                                           const BB_PointF* points,
                                           uint32_t filter);

/**
 * Decodes @p bytes straight into a texture, applying @p request on the way.
 *
 * The efficient path for an image that needs processing once and no more: it
 * never creates a CPU image, so nothing is retained beyond what Office holds.
 * Use BB_LoadImage instead when the same picture will be processed repeatedly.
 *
 * @p request may be NULL, which makes this BB_LoadTexture with extra steps.
 */
BB_API BB_Result BB_CALL BB_LoadTextureEx(const uint8_t* bytes,
                                          uint32_t length,
                                          const BB_ImageRequest* request,
                                          BB_Handle* out);

/** As BB_LoadTextureEx, reading the file itself. @p path is UTF-16. */
BB_API BB_Result BB_CALL BB_LoadTextureFromFileEx(const uint16_t* path,
                                                  const BB_ImageRequest* request,
                                                  BB_Handle* out);



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
                                              uint32_t count,
                                              uint32_t* applied);

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
 * Builds a texture from raw BGRA32 pixels, resampled to a chosen size.
 *
 * Separate from BB_LoadTexturePixels rather than an overload of it, so neither
 * call has ambiguous behaviour: this one always scales to exactly
 * @p targetWidth by @p targetHeight, using @p filter.
 *
 * @param pixels        BGRA32, @p height rows of @p stride bytes.
 * @param width         source pixels per row; must be greater than zero.
 * @param height        source rows; must be greater than zero.
 * @param stride        bytes per source row; at least `width * 4`.
 * @param targetWidth   destination pixels per row; must be greater than zero.
 * @param targetHeight  destination rows; must be greater than zero.
 * @param filter        one of the BB_SCALE_* values.
 * @param out           receives the texture handle.
 *
 * Both upscaling and downscaling are supported, at arbitrary ratios, and the
 * source stride may be padded. Alpha is interpolated correctly: the
 * interpolating filters work in premultiplied alpha internally, so a transparent
 * edge does not bleed dark colour into its visible neighbours.
 *
 * When the source and target sizes are equal the rows are copied and no filter
 * runs, so the result is bit-exact whichever filter was named.
 *
 * Every size and stride is validated, and every product is computed in 64 bits
 * and range-checked, before anything is read or allocated. Returns
 * BB_E_INVALID_ARG for a bad size, stride or filter, BB_E_OUT_OF_MEMORY if the
 * result cannot be allocated, and BB_E_DECODE_FAILED if the resampled pixels
 * cannot be turned into an image.
 *
 * ## What this does not control
 *
 * BlipBridge resamples the pixels **before** Office receives them. It does not
 * change how Office draws the resulting image. Do not read a choice of filter
 * here as control over PowerPoint's own rendering.
 */
BB_API BB_Result BB_CALL BB_LoadTexturePixelsScaled(const uint8_t* pixels,
                                                    uint32_t width,
                                                    uint32_t height,
                                                    int32_t stride,
                                                    uint32_t targetWidth,
                                                    uint32_t targetHeight,
                                                    uint32_t filter,
                                                    BB_Handle* out);

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
BB_API BB_Result BB_CALL BB_GetPictureCacheStats(uint32_t* textures,
                                                 uint32_t* shapes,
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
