#pragma once
/**
 * @file module_identity.hpp
 * What makes one loaded Office module distinguishable from another.
 *
 * ## Why a version string is not an identity
 *
 * A native profile is a set of byte offsets into one exact binary. "16.0.14334"
 * names a family of binaries, not a binary: the same marketing version ships
 * recompiled, and two modules can report the same `FileVersion` while differing
 * in every address a profile contains.
 *
 * So an identity here is the set of facts that actually distinguish an image:
 *
 *  - **architecture** - the PE machine field, because applying an x86 profile in
 *    an x64 process is the worst failure available to this library;
 *  - **version** - cheap to read and the first thing to reject on;
 *  - **PE timestamp and image size** - from the headers of the loaded image, and
 *    the pair that actually changes when Office is rebuilt;
 *  - **name**.
 *
 * The timestamp is `IMAGE_FILE_HEADER::TimeDateStamp`. On a deterministically
 * built module it is a content hash rather than a date, which suits this use
 * exactly: two images with the same value are the same image.
 *
 * ## Why this is also the cache key
 *
 * Resolution results are cached against an identity, never against a path and
 * never against an absolute address. A module that moves under ASLR is the same
 * module; a module that was patched is not, whatever it is called.
 */

#include "architecture.hpp"
#include <cstdint>
#include <string>

// MinGW requires the Windows base types first.
#include <windows.h>

namespace bb::native {

/// One loaded module, identified by what distinguishes it rather than by name.
struct ModuleIdentity {
    std::wstring name;
    Architecture architecture = Architecture::Unknown;

    /// From VS_FIXEDFILEINFO. Zero when the module carries no version resource.
    std::uint16_t major = 0;
    std::uint16_t minor = 0;
    std::uint16_t build = 0;
    std::uint16_t revision = 0;

    /// IMAGE_FILE_HEADER::TimeDateStamp - a content hash on a deterministic build.
    std::uint32_t timestamp = 0;
    /// IMAGE_OPTIONAL_HEADER::SizeOfImage, as mapped.
    std::uint32_t imageSize = 0;

    bool valid() const noexcept {
        return architecture != Architecture::Unknown && imageSize != 0;
    }

    /**
     * Whether two identities describe the same image.
     *
     * Every field participates. A comparison that ignored the timestamp would
     * accept a rebuilt module as the one a profile was derived from, which is
     * precisely the case this type exists to catch.
     */
    bool operator==(const ModuleIdentity& other) const noexcept {
        return architecture == other.architecture && major == other.major && minor == other.minor &&
               build == other.build && revision == other.revision && timestamp == other.timestamp &&
               imageSize == other.imageSize && name == other.name;
    }

    bool operator!=(const ModuleIdentity& other) const noexcept {
        return !(*this == other);
    }

    /// `oart.dll x86 16.0.14334.20906 ts=0x5F3A1B2C size=0xBC4000`, for logs and cache keys.
    std::wstring Describe() const;

    /// `16.0.14334.20906`, or an empty string when there is no version resource.
    std::wstring VersionText() const;
};

/**
 * Reads @p module's identity from the image mapped into this process.
 *
 * The version comes from the file on disk, because that is where the resource
 * is; everything else comes from the mapped headers, because that is what the
 * loader acted on. Returns an invalid identity rather than a partial one if any
 * of it cannot be read.
 */
ModuleIdentity IdentifyModule(HMODULE module) noexcept;

/// Convenience: identify an already-loaded module by name, without loading it.
ModuleIdentity IdentifyLoadedModule(const wchar_t* name) noexcept;

} // namespace bb::native
