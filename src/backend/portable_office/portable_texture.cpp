/**
 * @file portable_texture.cpp
 * Implementation of the portable texture resource. See the header for the
 * resource model and how it differs from the accelerated one.
 */

#include "portable_texture.hpp"

#include "../../image/decode.hpp"
#include "../../image/encode.hpp"
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstring>
#include <map>
#include <sstream>

// MinGW requires the Windows base types before anything else Windows.
#include <windows.h>

namespace bb::portable {
namespace {

/**
 * Where texture files go.
 *
 * Per-process, because two PowerPoint instances with the add-in loaded would
 * otherwise write into each other's directory and one exiting would delete
 * files the other was still using. The process id makes that impossible without
 * any locking.
 */
const std::wstring& TemporaryDirectory() {
    static const std::wstring directory = [] {
        wchar_t buffer[MAX_PATH + 1] = {};
        const DWORD length = GetTempPathW(MAX_PATH, buffer);
        std::wstring base =
            (length > 0 && length <= MAX_PATH) ? std::wstring(buffer, length) : std::wstring(L".");
        if (!base.empty() && base.back() != L'\\') {
            base.push_back(L'\\');
        }
        std::wostringstream out;
        out << base << L"BlipBridge-" << GetCurrentProcessId() << L"\\";
        std::wstring path = out.str();
        // Failure here is not fatal yet: the first write will report it, and it
        // will report it with the path, which is far more use than a failure at
        // library load with no context.
        CreateDirectoryW(path.c_str(), nullptr);
        return path;
    }();
    return directory;
}

/// Narrows a wide path for an error message. Diagnostics only.
std::string Narrow(const std::wstring& text) {
    const int size =
        WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (size <= 1) {
        return std::string();
    }
    std::string out(static_cast<std::size_t>(size - 1), ' ');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, out.data(), size, nullptr, nullptr);
    return out;
}

/// Builds the message for a file operation that failed, naming the file.
std::string DescribeFileFailure(const char* what, const std::wstring& path, DWORD code) {
    std::ostringstream out;
    out << "BlipBridge could not " << what << " the temporary image file " << Narrow(path)
        << " (Windows error " << code
        << "). The portable backend needs a file because Office's Fill.UserPicture takes a path";
    return out.str();
}

/**
 * Writes @p bytes to @p path, replacing whatever was there.
 *
 * FILE_SHARE_READ so that Office can read the file while this handle is still
 * open. It will not need to, but a sharing violation in the middle of a fill
 * would be a miserable thing to debug.
 */
void WriteWholeFile(const std::wstring& path, const std::vector<std::uint8_t>& bytes) {
    const auto open = [&path] {
        return CreateFileW(path.c_str(),
                           GENERIC_WRITE,
                           FILE_SHARE_READ,
                           nullptr,
                           CREATE_ALWAYS,
                           FILE_ATTRIBUTE_TEMPORARY,
                           nullptr);
    };

    HANDLE file = open();
    if (file == INVALID_HANDLE_VALUE && GetLastError() == ERROR_PATH_NOT_FOUND) {
        /*
         * The process directory is gone, and the usual reason is that we removed
         * it ourselves: clearing every texture takes the directory down with the
         * files. A library that has been cleared has to keep working afterwards,
         * so put it back and try once more rather than failing a perfectly
         * ordinary apply. Creating it is also cheap enough that checking first
         * would cost more than retrying.
         */
        CreateDirectoryW(TemporaryDirectory().c_str(), nullptr);
        file = open();
    }
    if (file == INVALID_HANDLE_VALUE) {
        const DWORD code = GetLastError();
        throw bb::Error(HRESULT_FROM_WIN32(code), DescribeFileFailure("create", path, code));
    }

    struct Close {
        HANDLE value;

        ~Close() {
            CloseHandle(value);
        }
    } close{file};

    std::size_t written = 0;
    while (written < bytes.size()) {
        const std::size_t remaining = bytes.size() - written;
        const DWORD chunk = static_cast<DWORD>(remaining > 0x10000000u ? 0x10000000u : remaining);
        DWORD done = 0;
        if (!WriteFile(file, bytes.data() + written, chunk, &done, nullptr) || done == 0) {
            const DWORD code = GetLastError();
            throw bb::Error(HRESULT_FROM_WIN32(code), DescribeFileFailure("write", path, code));
        }
        written += done;
    }
}

/**
 * The public handle table.
 *
 * Handles start at the same base as the accelerated store so that a handle means
 * the same kind of thing on both backends, and stay well below the image store's
 * `0x4000000` so the ABI can keep telling the two apart by value.
 */
constexpr std::uint64_t kHandleBase = 0x1000000ull;

class HandleTable {
  public:
    static HandleTable& Instance() {
        static HandleTable table;
        return table;
    }

    std::uint64_t Add(TextureRef texture) {
        const std::uint64_t handle = next_++;
        textures_.emplace(handle, std::move(texture));
        return handle;
    }

    TextureRef Find(std::uint64_t handle) const {
        const auto entry = textures_.find(handle);
        return entry == textures_.end() ? TextureRef() : entry->second;
    }

    /**
     * True when this handle was ever handed out, live or not.
     *
     * Distinct from Find on purpose. A handle that was issued and released is a
     * different thing from one that was never created, and telling them apart
     * is what lets the refusal say so - which is the whole visible consequence
     * of handles never being recycled. Since the counter only ever rises, the
     * range below is an exact answer rather than a heuristic.
     */
    bool EverIssued(std::uint64_t handle) const noexcept {
        return handle >= kHandleBase && handle < next_;
    }

    bool Remove(std::uint64_t handle) {
        return textures_.erase(handle) != 0;
    }

    void Clear() noexcept {
        textures_.clear();
    }

    std::size_t Count() const noexcept {
        return textures_.size();
    }

    /// The handle the next registration will get. Diagnostics only.
    std::uint64_t NextHandle() const noexcept {
        return next_;
    }

  private:
    std::map<std::uint64_t, TextureRef> textures_;
    // Never rewound, so a released handle is never handed out again.
    std::uint64_t next_ = kHandleBase;
};

/// Office edits made by this process. See OfficeFillCount in the header.
std::uint64_t g_officeFills = 0;

/// Images built since the process started. Diagnostics only; never rewound.
std::uint64_t g_creations = 0;

} // namespace

std::uint64_t OfficeFillCount() noexcept {
    return g_officeFills;
}

void CountOfficeFill() noexcept {
    ++g_officeFills;
}

std::uint64_t NextSharedImageId() noexcept {
    // Process-unique, never reused. Zero is reserved for "no image".
    static std::uint64_t next = 1;
    return next++;
}

PortableTexture::PortableTexture(std::vector<std::uint8_t> pixels,
                                 std::uint32_t width,
                                 std::uint32_t height,
                                 std::uint64_t id)
    : pixels_(std::move(pixels)), width_(width), height_(height), id_(id) {
    ++g_creations;
}

PortableTexture::~PortableTexture() {
    if (!path_.empty()) {
        // Best effort: a file that will not delete is a stray temporary, which
        // is not worth failing a destructor over.
        DeleteFileW(path_.c_str());
    }
}

const std::wstring& PortableTexture::Path() {
    if (!path_.empty()) {
        return path_;
    }

    std::vector<std::uint8_t> encoded;
    const image::EncodeStatus status = image::EncodePng(
        pixels_.data(), width_, height_, static_cast<std::int32_t>(width_) * 4, encoded);
    if (status != image::EncodeStatus::Ok) {
        throw bb::Error(E_FAIL, image::DescribeEncodeStatus(status));
    }

    std::wostringstream name;
    name << TemporaryDirectory() << L"texture-" << id_ << L".png";
    const std::wstring path = name.str();
    WriteWholeFile(path, encoded);
    // Assigned only once the file is actually on disk, so a failed write leaves
    // the texture as it was and the next apply tries again rather than handing
    // Office the path of a file that is not there.
    path_ = path;
    return path_;
}

TextureRef CreateTextureFromBytes(const std::uint8_t* bytes, std::size_t length) {
    if (!bytes || length == 0) {
        throw bb::Error(E_INVALIDARG, "Image bytes are required");
    }
    image::DecodedImage decoded;
    const image::DecodeStatus status = image::Decode(bytes, length, decoded);
    if (status != image::DecodeStatus::Ok) {
        throw bb::Error(BB_E_INVALID_IMAGE, image::DescribeDecodeStatus(status));
    }
    return std::make_shared<PortableTexture>(
        std::move(decoded.pixels), decoded.width, decoded.height, NextSharedImageId());
}

TextureRef CreateTextureFromPixels(const void* pixels,
                                   std::uint32_t width,
                                   std::uint32_t height,
                                   std::int32_t stride) {
    if (!pixels || width == 0 || height == 0) {
        throw bb::Error(E_INVALIDARG, "Pixels and non-zero dimensions are required");
    }
    const std::int64_t row = static_cast<std::int64_t>(width) * 4;
    if (static_cast<std::int64_t>(stride) < row) {
        throw bb::Error(E_INVALIDARG, "Stride must be at least width*4 bytes for BGRA32");
    }

    // Copied tightly packed, which is what everything downstream expects and
    // what makes the stride question disappear after this point.
    std::vector<std::uint8_t> copy(static_cast<std::size_t>(row) * height);
    const auto* source = static_cast<const std::uint8_t*>(pixels);
    for (std::uint32_t y = 0; y < height; ++y) {
        std::memcpy(copy.data() + static_cast<std::size_t>(row) * y,
                    source + static_cast<std::size_t>(stride) * y,
                    static_cast<std::size_t>(row));
    }
    return std::make_shared<PortableTexture>(std::move(copy), width, height, NextSharedImageId());
}

std::uint64_t RegisterTexture(TextureRef texture) {
    if (!texture) {
        throw bb::Error(E_INVALIDARG, "No texture to register");
    }
    return HandleTable::Instance().Add(std::move(texture));
}

TextureRef LookupTexture(std::uint64_t handle) {
    TextureRef texture = HandleTable::Instance().Find(handle);
    if (!texture) {
        std::ostringstream out;
        out << "Texture handle " << handle << " is not valid (";
        out << (HandleTable::Instance().EverIssued(handle)
                    ? "it was released; handles are never recycled"
                    : "it was never created");
        out << ")";
        throw bb::Error(BB_E_TEXTURE_NOT_FOUND, out.str());
    }
    return texture;
}

bool IsTextureHandle(std::uint64_t handle) noexcept {
    return handle >= kHandleBase;
}

bool OwnsHandle(std::uint64_t handle) noexcept {
    // Issued, not necessarily still live: a caller asking whose handle this is
    // needs the same answer before and after a release, or a released texture
    // handle would start looking like somebody else's.
    return HandleTable::Instance().EverIssued(handle);
}

std::uint64_t TextureIdOf(const TextureRef& texture) {
    return texture ? texture->id() : 0;
}

void ReleaseTexture(std::uint64_t handle) {
    if (!HandleTable::Instance().Remove(handle)) {
        std::ostringstream out;
        out << "Texture handle " << handle << " is not valid (";
        out << (HandleTable::Instance().EverIssued(handle)
                    ? "it was already released; handles are never recycled"
                    : "it was never created");
        out << ")";
        throw bb::Error(BB_E_TEXTURE_NOT_FOUND, out.str());
    }
}

void ClearTextures() noexcept {
    HandleTable::Instance().Clear();
}

std::size_t TextureCount() noexcept {
    return HandleTable::Instance().Count();
}

std::wstring DescribeTextures(std::uint64_t handle) noexcept {
    try {
        HandleTable& table = HandleTable::Instance();
        std::wostringstream out;
        out << L"textures=" << table.Count() << L";nextHandle=" << table.NextHandle()
            << L";creations=" << g_creations << L";fills=" << g_officeFills << L';';
        if (handle > 0) {
            const TextureRef texture = table.Find(handle);
            out << L"handle=" << handle << L';';
            if (texture) {
                out << L"width=" << texture->width() << L";height=" << texture->height()
                    << L";bytes=" << texture->byteCount() << L";applies="
                    << texture->applyCount()
                    // Whether the PNG has been written yet. Zero here after an
                    // apply would mean the fill never needed the file, which
                    // cannot happen - so it is worth being able to see.
                    << L";encoded=" << (texture->encoded() ? 1 : 0) << L';';
            } else {
                out << L"live=0;released=" << (table.EverIssued(handle) ? 1 : 0) << L';';
            }
        }
        return out.str();
    } catch (...) {
        return L"error=1;";
    }
}

void CleanUpTemporaryDirectory() noexcept {
    // Only ever removes an empty directory, so a file that outlived its texture
    // is left where a human can find it rather than being deleted blind.
    RemoveDirectoryW(TemporaryDirectory().c_str());
}

} // namespace bb::portable
