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
#include <cstring>
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

    /**
     * The CodeView build signature: the PDB GUID and age from the debug directory.
     *
     * This is the strong part of the identity, and it is why version, timestamp
     * and image size are not relied on alone to authorise executing private
     * Office internals. The linker generates a fresh GUID for every build, so two
     * images sharing one are the same build in a way three coincidable integers
     * cannot establish.
     *
     * It is also nearly free to read - twenty bytes from a read-only directory -
     * which matters because it is checked on every process start. A full file
     * hash of ppcore.dll, oart.dll and gfx.dll is about 38 MB of reading; see
     * validation_cache.hpp for where that cost is actually paid.
     *
     * `hasBuildSignature` false means the module carries no debug directory. That
     * is a fact about the module rather than a failure, but it does mean the
     * identity is not strong enough to authorise private calls.
     */
    bool hasBuildSignature = false;
    std::uint8_t buildGuid[16] = {};
    std::uint32_t buildAge = 0;

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
        if (architecture != other.architecture || major != other.major || minor != other.minor ||
            build != other.build || revision != other.revision || timestamp != other.timestamp ||
            imageSize != other.imageSize || name != other.name) {
            return false;
        }
        // The build signature participates whenever either side has one. A
        // profile derived against a signed build must not match an image with
        // none, and two images with different GUIDs are different builds whatever
        // their timestamps say.
        if (hasBuildSignature != other.hasBuildSignature) {
            return false;
        }
        if (!hasBuildSignature) {
            return true;
        }
        return buildAge == other.buildAge &&
               std::memcmp(buildGuid, other.buildGuid, sizeof(buildGuid)) == 0;
    }

    bool operator!=(const ModuleIdentity& other) const noexcept {
        return !(*this == other);
    }

    /// `oart.dll x86 16.0.14334.20906 ts=0x5F3A1B2C size=0xBC4000`, for logs and cache keys.
    std::wstring Describe() const;

    /// `16.0.14334.20906`, or an empty string when there is no version resource.
    std::wstring VersionText() const;

    /// The build GUID and age, or an empty string when the module carries none.
    std::wstring BuildSignatureText() const;

    /**
     * Whether this identity may authorise calling private Office internals.
     *
     * Version, timestamp and image size identify a build well enough to look one
     * up. They are not what a decision to execute reverse-engineered code should
     * rest on, so that decision additionally requires a build signature.
     */
    bool strong() const noexcept {
        return valid() && hasBuildSignature;
    }
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
