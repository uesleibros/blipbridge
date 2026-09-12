#pragma once
/**
 * @file memory_safety.hpp
 * The guarded reads every other file here is built on.
 *
 * Nothing in the native backend may dereference a pointer it has not first
 * established is committed and readable. The pointers involved come from walking
 * Office's own object graph, and a graph that has changed shape between builds
 * produces pointers that are plausible and wrong - so "it worked last version"
 * is not a reason to read something.
 *
 * These are also what makes the read-only research probes safe to run inside a
 * user's PowerPoint: the worst outcome of a wrong guess is a refusal.
 *
 * Architecture-neutral. Pointer size comes from the compiler.
 */

#include <cstddef>
#include <cstdint>

// MinGW requires the Windows base types first.
#include <windows.h>

namespace bb::native {

/// True when every page spanning [address, address+size) is committed and readable.
bool IsReadable(const void* address, std::size_t size) noexcept;

/**
 * True when @p address is committed and lies in an executable region.
 *
 * A vtable slot must satisfy this. A slot pointing into a data section is the
 * clearest possible sign that what was read is not a vtable, and it is worth
 * catching there rather than at the call.
 */
bool IsExecutable(const void* address) noexcept;

/**
 * Reads one pointer-sized word at @p base + @p offset, or 0 when it is not
 * readable.
 *
 * Zero is used as the failure value deliberately: every use of this in the
 * native backend goes on to require the result to be a valid pointer into an
 * expected module, and 0 fails that check. There is no valid object at 0 to
 * confuse it with.
 */
std::uintptr_t ReadPointer(const void* base, std::size_t offset) noexcept;

/// Reads one 32-bit word, or 0 when it is not readable.
std::uint32_t ReadDword(const void* base, std::size_t offset) noexcept;

/// Compares @p size bytes at @p address against @p expected, guarded.
bool BytesMatch(const void* address, const std::uint8_t* expected, std::size_t size) noexcept;

} // namespace bb::native
