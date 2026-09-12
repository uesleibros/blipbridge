#pragma once
/**
 * @file architecture.hpp
 * Which architecture this build is, and which one a module is.
 *
 * ## Why this is not a one-line answer
 *
 * There are two different questions and conflating them is how a profile gets
 * applied to the wrong binary:
 *
 *  - **What am I?** A compile-time fact. The build either contains the x64
 *    native backend or the x86 one, never both, and `Current()` says which.
 *  - **What is that module?** A runtime fact, read from the module's PE header.
 *
 * A native profile is a set of addresses into a specific binary. Applying an x86
 * profile inside an x64 process - or the reverse - means reading and eventually
 * calling addresses that mean nothing, which is the single worst failure
 * available to this library. So a profile carries its architecture, a module
 * reports its architecture, and the two are compared before any offset is used.
 *
 * The comparison is against the **PE machine field**, never against a file name
 * or a directory. `Program Files (x86)` is a convention; `IMAGE_FILE_MACHINE_I386`
 * is what the loader acted on.
 */

#include <cstdint>

// MinGW requires the Windows base types first.
#include <windows.h>

namespace bb::native {

/// The architectures BlipBridge has, or could have, a native backend for.
enum class Architecture {
    Unknown,
    X86,
    X64,
};

const char* Name(Architecture architecture) noexcept;

/**
 * The architecture this build targets, decided by the compiler.
 *
 * Deliberately not a runtime probe. A build contains one native backend, chosen
 * in CMake, and code that asked "which am I" at runtime would be code that could
 * be wrong about it.
 */
constexpr Architecture Current() noexcept {
#if defined(_WIN64)
    return Architecture::X64;
#elif defined(_WIN32)
    return Architecture::X86;
#else
    return Architecture::Unknown;
#endif
}

/**
 * Reads @p module's PE machine field.
 *
 * Returns Unknown for anything unreadable or unrecognised rather than guessing.
 * A module whose architecture cannot be established is one no profile may be
 * applied to.
 */
Architecture ArchitectureOf(HMODULE module) noexcept;

/// True when @p module is the architecture this build has a backend for.
bool MatchesCurrent(HMODULE module) noexcept;

} // namespace bb::native
