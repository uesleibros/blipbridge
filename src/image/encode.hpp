#pragma once
/**
 * @file encode.hpp
 * Encoding canonical BGRA32 back to PNG, through Windows Imaging Component.
 *
 * ## Why this exists
 *
 * The accelerated backend never needs it: it hands decoded pixels straight to
 * Office's own image cache and no encoded form is ever produced. The portable
 * backend cannot do that. `Fill.UserPicture` takes a **path**, so a texture that
 * began life as raw pixels - a resample, a crop, a quad warp - has to become a
 * file before Office will look at it.
 *
 * That is the whole reason this file exists, and it is worth being explicit
 * about: encoding is a cost the portable path pays and the native path does not.
 *
 * ## Why PNG, and not JPEG
 *
 * Lossless and alpha-capable. A texture that has been through the pipeline is
 * the caller's image, and handing Office something visibly different from what
 * they asked for - JPEG ringing along every hard edge, an alpha channel silently
 * flattened onto black - would be a defect that is very hard to attribute later.
 * The file is a transport detail between this library and Office; it is written
 * to a temporary directory and deleted again, so its size matters much less than
 * its fidelity.
 *
 * ## The format contract
 *
 * Input is what the rest of `src/image` produces and consumes: **32bppBGRA,
 * straight (non-premultiplied) alpha, top-down, tightly packed or with a stride
 * at least `width * 4`**. WIC's `GUID_WICPixelFormat32bppBGRA` is straight alpha,
 * so the pixels are handed over unchanged - no premultiply, no conversion, no
 * opportunity to darken a transparent edge.
 */

// MinGW needs the Windows base types before anything COM.
#include <windows.h>

#include <cstdint>
#include <vector>

namespace bb::image {

/// Why an encode was refused. Each value says something a caller can act on.
enum class EncodeStatus {
    Ok,
    /// No pixels, a zero dimension, or a stride narrower than a row.
    InvalidArgument,
    /// WIC could not be created at all - a broken or locked-down installation.
    NoEncoder,
    /// More pixels than `kMaxPixels`, the same ceiling the resampler enforces.
    UnsupportedSize,
    /// WIC accepted the request and then failed part-way through it.
    EncodeFailed,
    OutOfMemory,
};

/// A sentence explaining a status, suitable for a caller's error message.
const char* DescribeEncodeStatus(EncodeStatus status) noexcept;

/**
 * Encodes BGRA32 pixels to PNG bytes.
 *
 * @param pixels  BGRA32, straight alpha, top-down.
 * @param width   in pixels, non-zero.
 * @param height  in pixels, non-zero.
 * @param stride  bytes per row; may exceed `width * 4`.
 * @param out     receives the PNG file's bytes; untouched unless this returns Ok.
 *
 * Never throws. Encoding to memory rather than straight to a file is deliberate:
 * it keeps this testable without a filesystem, and it lets the caller decide
 * where the bytes go and how failures to write one are reported.
 */
EncodeStatus EncodePng(const std::uint8_t* pixels,
                       std::uint32_t width,
                       std::uint32_t height,
                       std::int32_t stride,
                       std::vector<std::uint8_t>& out) noexcept;

} // namespace bb::image
