#pragma once
/**
 * @file decode.hpp
 * Decoding encoded images to canonical BGRA32, through Windows Imaging
 * Component.
 *
 * ## Why this exists
 *
 * `BB_LoadTexture` already hands encoded bytes to Office's own image pipeline,
 * which decodes them. But the scaling API works on decoded BGRA, so a caller who
 * wanted an encoded image resampled had to decode it themselves first - in VBA,
 * through GDI+, into a `Byte()` - and hand back pixels. That is a lot of
 * marshalling to arrive where this file starts.
 *
 * So: decode here, in canonical BGRA32, and hand the result to the resampler
 * that already exists. No filter is reimplemented; `resample.hpp` remains the one
 * place any of that maths lives.
 *
 * ## Why WIC
 *
 * It ships with Windows, it is the decoder Office itself is built over, and it
 * converts to a requested pixel format as part of the decode rather than making
 * this file walk source formats. No third-party image library is worth a
 * dependency for work the operating system already does.
 *
 * ## The canonical format, chosen once
 *
 * **32bppBGRA, straight (non-premultiplied) alpha, top-down, tightly packed.**
 *
 * That is what `resample.hpp` documents as its input, what
 * `CreateCachedImageFromPixels` takes, and what the existing raw-pixel entry
 * points already accept, so a decoded image and a caller-supplied buffer are the
 * same thing from here on. Straight alpha specifically: the resampler converts
 * to premultiplied internally where it needs to, and doing it twice would darken
 * every transparent edge.
 *
 * Formats without alpha decode to opaque BGRA with the alpha byte set to 255 by
 * the format conversion, so a JPEG and a PNG arrive here indistinguishable.
 *
 * ## What is tested, and what is merely likely
 *
 * PNG (with and without alpha), JPEG and BMP are tested against generated
 * fixtures. WIC decodes a good deal more - GIF, TIFF, ICO, JPEG XR, and whatever
 * codecs are installed on the machine - and those will very probably work, but
 * an untested format is not a supported one and this header does not claim them.
 */

// MinGW needs the Windows base types before anything COM.
#include <windows.h>

#include <cstdint>
#include <vector>

namespace bb::image {

/// Why a decode was refused. Each value says something a caller can act on.
enum class DecodeStatus {
    Ok,
    /// No bytes, or a zero length.
    NoData,
    /// WIC could not be created at all - a broken or locked-down installation.
    NoDecoder,
    /// The bytes are not an image WIC recognises.
    UnknownFormat,
    /// Recognised, but the data is damaged or ends early.
    Corrupt,
    /// Decoded, but the image has no pixels, or more than kMaxPixels of them.
    UnsupportedSize,
    /// The conversion to BGRA32 failed.
    ConversionFailed,
    OutOfMemory,
};

/// A sentence explaining a status, suitable for a caller's error message.
const char* DescribeDecodeStatus(DecodeStatus status) noexcept;

/// One decoded image: tightly packed BGRA32, `width * height * 4` bytes.
struct DecodedImage {
    std::uint32_t width = 0;
    std::uint32_t height = 0;
    std::vector<std::uint8_t> pixels;

    /// Bytes per row. Always tight - there is no padding to reason about.
    std::int32_t stride() const noexcept {
        return static_cast<std::int32_t>(width) * 4;
    }
};

/**
 * Decodes @p bytes into BGRA32.
 *
 * @param bytes   encoded image data - PNG, JPEG or BMP among others.
 * @param length  its length in bytes.
 * @param out     receives the decoded image; untouched unless this returns Ok.
 *
 * Never throws and never reads past @p length. Malformed, truncated and
 * hostile input produce a status, not a crash: WIC does the parsing, and every
 * size it reports is range-checked here before an allocation is made from it.
 *
 * The first frame is used. Multi-frame formats decode their first frame rather
 * than being refused, which is what a caller asking for "the image" means.
 */
DecodeStatus Decode(const std::uint8_t* bytes, std::size_t length, DecodedImage& out) noexcept;

/**
 * Reads @p path and decodes it.
 *
 * Exists because the alternative is every caller writing the same file read, and
 * a VBA caller writing it badly. The file is read whole before decoding - these
 * are slide textures, not a streaming workload - and a file larger than
 * `kMaxEncodedBytes` is refused rather than read.
 *
 * @p path is UTF-16, because that is what the platform's file APIs take.
 */
DecodeStatus DecodeFile(const wchar_t* path, DecodedImage& out) noexcept;

/**
 * The largest encoded file this will read, in bytes.
 *
 * A decoded 64-megapixel image is already the resampler's ceiling at 256 MB;
 * an encoded file that size is either a mistake or an attack, and refusing it
 * before the read is cheaper than finding out afterwards.
 */
inline constexpr std::uint64_t kMaxEncodedBytes = 256ull * 1024ull * 1024ull;

} // namespace bb::image
