#pragma once
/**
 * @file validation_cache.hpp
 * Remembers that a native profile once proved itself, so it need not prove
 * itself on every start.
 *
 * ## What is cached, and what is not
 *
 * **Cached:** "the modules with these exact identities passed the deterministic
 * self-test, under resolver version N". That is a statement about a set of
 * binaries, and it stays true for as long as those binaries do.
 *
 * **Never cached:** any address. Not a resolved RVA, and above all not a process
 * virtual address - ASLR makes those meaningless across runs, and a cache that
 * handed back a stale absolute pointer would be the most dangerous thing in this
 * library. Resolution is redone each process; only the *verdict* is remembered.
 *
 * ## Why the verdict is worth caching when the addresses are not
 *
 * The self-test is the expensive part and the part with side effects: it creates
 * a disposable Shape, applies a known image through reverse-engineered internals,
 * and checks the result. Doing that on every PowerPoint start would be a visible
 * cost for a question whose answer cannot have changed while the binaries have
 * not.
 *
 * Structural validation, by contrast, is cheap and is *not* skipped. Every
 * process re-checks that the facts still describe the memory in front of it.
 * The cache removes a proof, never a check.
 *
 * ## Invalidation
 *
 * The key is every module's full identity, build signature included. Any change
 * to any of them - an Office update, a repaired install, a patched image -
 * produces a different key and therefore a miss, and a miss means the self-test
 * runs again. There is no eviction policy because there is nothing to evict: a
 * stale entry is simply never looked up.
 *
 * A recorded *failure* is honoured for the life of the process but not written
 * to the store. A failed self-test may be a transient host problem, and
 * condemning a build permanently on one bad run would be worse than re-testing
 * it next time.
 *
 * ## Where it lives
 *
 * `HKCU\Software\BlipBridge\NativeValidation`. Per-user, no administrator, and
 * no file for another process to race on. Nothing in it can direct BlipBridge at
 * an address: the worst a tampered entry can do is skip a self-test for a module
 * set that must still pass structural validation, which is why the entries are
 * verdicts rather than data.
 */

#include "profile.hpp"
#include <string>

namespace bb::native {

/// The self-test outcome recorded against one set of module identities.
enum class CachedVerdict {
    /// No entry: the self-test has never been run for these exact binaries.
    Unknown,
    /// It passed. The self-test may be skipped; structural validation still runs.
    Passed,
};

/**
 * Bumped whenever the meaning of a stored verdict changes.
 *
 * A cached "passed" is only as good as the self-test that produced it. If the
 * self-test is strengthened, or the facts a profile carries change shape, every
 * stored verdict was reached under different rules and must be discarded -
 * which this does simply by becoming part of the key.
 */
inline constexpr std::uint32_t kValidationCacheVersion = 1;

/// The key for one profile: every module identity it depends on, in order.
std::wstring ValidationKey(const NativeProfile& profile);

/// What is recorded for @p profile's modules, if anything.
CachedVerdict LookUpValidation(const NativeProfile& profile) noexcept;

/// Records that @p profile's self-test passed. Failures are deliberately not stored.
void RecordValidationPassed(const NativeProfile& profile) noexcept;

/// Forgets everything. Tests and diagnostics only.
void ForgetValidationCache() noexcept;

} // namespace bb::native
