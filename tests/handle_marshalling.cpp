/**
 * @file handle_marshalling.cpp
 * Proves the 64-bit texture handle survives the VBA boundary on both architectures.
 *
 * ## What this can and cannot establish
 *
 * The public VBA type is a pair of signed Longs:
 *
 *     Public Type BlipBridgeTexture
 *         Low As Long
 *         High As Long
 *     End Type
 *
 * because `LongLong` does not exist in 32-bit VBA and a `Double` loses precision
 * above 2^53. This file reproduces the wrapper's arithmetic in C++ and checks it
 * against the values it must carry, so a sign-extension or truncation bug fails
 * here rather than silently applying the wrong texture in a document.
 *
 * It does **not** run VBA. A real 32-bit PowerPoint executing the wrapper is a
 * separate, still-unvalidated thing, and nothing here should be read as covering
 * it. What this covers is the arithmetic and the slot layout, on the
 * architecture it is compiled for.
 */

#include <blipbridge/blipbridge.h>

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>

namespace {

int g_failures = 0;

void Check(bool condition, const std::string& what) {
    if (condition) {
        std::printf("  ok   %s\n", what.c_str());
        return;
    }
    std::printf("  FAIL %s\n", what.c_str());
    ++g_failures;
}

/// The wrapper's public type, laid out exactly as VBA lays it out.
struct VbaTexture {
    std::int32_t low = 0;
    std::int32_t high = 0;
};

/**
 * What the wrapper receives from an out-parameter.
 *
 * VBA passes the struct ByRef, so the ABI writes a uint64_t straight into it.
 * That is a reinterpretation, not a conversion - which is the point, and why
 * there is no arithmetic on this path at all.
 */
VbaTexture FromNative(std::uint64_t handle) {
    VbaTexture texture;
    static_assert(sizeof(VbaTexture) == sizeof(std::uint64_t),
                  "the VBA type must be exactly eight bytes");
    std::memcpy(&texture, &handle, sizeof(handle));
    return texture;
}

/**
 * The x64 wrapper's TextureToNative, transcribed.
 *
 * `Low` is a signed Long, so masking to 32 bits before combining is what stops a
 * handle above 0x7FFFFFFF sign-extending into the high half.
 */
std::uint64_t ToNativeX64(const VbaTexture& texture) {
    const std::uint64_t low = static_cast<std::uint64_t>(texture.low) & 0xFFFFFFFFull;
    const std::uint64_t high = static_cast<std::uint64_t>(texture.high) & 0xFFFFFFFFull;
    return (high * 0x100000000ull) | low;
}

/**
 * The x86 wrapper passes Low then High as two stack slots.
 *
 * A stdcall frame carries a 64-bit value as exactly that, low word first, so
 * reassembling them in that order is what the callee sees.
 */
std::uint64_t ToNativeX86(const VbaTexture& texture) {
    const std::uint32_t low = static_cast<std::uint32_t>(texture.low);
    const std::uint32_t high = static_cast<std::uint32_t>(texture.high);
    return (static_cast<std::uint64_t>(high) << 32) | low;
}

std::string Hex(std::uint64_t value) {
    char buffer[32];
    std::snprintf(buffer, sizeof(buffer), "0x%016llX",
                  static_cast<unsigned long long>(value));
    return buffer;
}

} // namespace

int main() {
    /*
     * The values that matter: zero, one, the 32-bit boundary from both sides, a
     * pattern with every nibble distinct, and each sign boundary. Anything that
     * sign-extends, truncates, or routes through a double fails on at least one
     * of these.
     */
    const std::uint64_t values[] = {
        0x0000000000000000ull,
        0x0000000000000001ull,
        0x00000000FFFFFFFFull,
        0x0000000100000000ull,
        0x0123456789ABCDEFull,
        0x7FFFFFFFFFFFFFFFull,
        0x8000000000000000ull,
        0xFFFFFFFFFFFFFFFFull,
        // The handle space this library actually issues starts here.
        0x0000000001000000ull,
    };

    Check(sizeof(BB_Handle) == 8, "BB_Handle is 64 bits");
    Check(sizeof(VbaTexture) == 8, "the VBA texture type is 64 bits");

    for (const std::uint64_t value : values) {
        const VbaTexture texture = FromNative(value);

        const std::uint64_t viaX64 = ToNativeX64(texture);
        Check(viaX64 == value, "x64 reconstruction is bit-exact for " + Hex(value));

        const std::uint64_t viaX86 = ToNativeX86(texture);
        Check(viaX86 == value, "x86 slot layout is bit-exact for " + Hex(value));

        // The two architectures must agree, or the same handle would mean
        // different things depending on which PowerPoint opened the file.
        Check(viaX64 == viaX86, "both architectures agree on " + Hex(value));
    }

    // Sign behaviour, stated rather than assumed: the top half of the 32-bit
    // range reads back as a negative Long, and that is the correct bit pattern.
    {
        const VbaTexture texture = FromNative(0x00000000FFFFFFFFull);
        Check(texture.low == -1, "a low word of 0xFFFFFFFF reads as -1 in a signed Long");
        Check(texture.high == 0, "and the high word stays zero");
        Check(ToNativeX64(texture) == 0x00000000FFFFFFFFull,
              "and it still reassembles to 0x00000000FFFFFFFF");
    }
    {
        const VbaTexture texture = FromNative(0xFFFFFFFFFFFFFFFFull);
        Check(texture.low == -1 && texture.high == -1,
              "an all-ones handle is two -1 Longs");
        Check(ToNativeX64(texture) == 0xFFFFFFFFFFFFFFFFull,
              "and reassembles without loss");
    }

    /*
     * The reason the public type is not a Double. 2^53 + 1 is the first integer a
     * double cannot represent, and a handle carried through one would come back
     * as its neighbour - applying or releasing the wrong texture, silently.
     */
    {
        const std::uint64_t beyondDouble = (1ull << 53) + 1ull;
        const VbaTexture texture = FromNative(beyondDouble);
        Check(ToNativeX64(texture) == beyondDouble,
              "a value a double could not hold survives the Long pair exactly");
        const double asDouble = static_cast<double>(beyondDouble);
        Check(static_cast<std::uint64_t>(asDouble) != beyondDouble,
              "and the same value really would be lost by a double");
    }

    if (g_failures != 0) {
        std::printf("\n%d handle marshalling checks failed\n", g_failures);
        return 1;
    }
    std::printf("\nall handle marshalling checks passed\n");
    std::printf("Note: this validates the arithmetic and slot layout, not a real\n");
    std::printf("32-bit PowerPoint executing the VBA wrapper.\n");
    return 0;
}
