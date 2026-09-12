/**
 * @file chain_probe.cpp
 * Follows a pointer chain from `Shape.Fill` and reports what it finds, without
 * calling any of it.
 *
 * ## Why a chain rather than one object
 *
 * `structure_probe.cpp` answers "what is the Fill wrapper". The next questions
 * are about what hangs off it, and they cannot be answered one probe at a time:
 * identifying the object at `+0x4` means looking at *its* vtable, *its* fields,
 * and what *they* point at. So the walk is driven from the harness - a chain of
 * offsets - and the probe reports each hop.
 *
 * Every hop is a guarded read and nothing is ever called. A wrong offset yields
 * "not readable", which is the intended failure.
 *
 * ## Why it reports a digest as well as the bytes
 *
 * The strongest evidence available for "what is this object for" is what changes
 * in it when Office does something. A harness can dump the chain, apply a picture
 * fill through the *documented* API, dump again, and compare - and the fields
 * that moved are the fields that hold the fill. That is evidence about role,
 * obtained without executing anything private.
 *
 * The digest makes that comparison cheap and exact: a stable hash of the object's
 * leading words, so a harness can tell "nothing changed" from "something did"
 * before reading any of it.
 *
 * Research only.
 */

#include "../../src/backend/native/common/memory_safety.hpp"
#include "../../src/backend/native/common/structural_validation.hpp"
#include "../experiment_api.hpp"
#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>
#include <cstdint>
#include <iomanip>
#include <sstream>
#include <string>
#include <vector>

// MinGW requires the Windows base types before the Automation declarations.
#include <windows.h>

namespace {

/// The Office modules a walk from Shape.Fill may legitimately land in.
bb::native::ModuleWorld OfficeWorld() {
    bb::native::ModuleWorld world;
    for (const wchar_t* name : {L"POWERPNT.EXE",
                                L"ppcore.dll",
                                L"oart.dll",
                                L"gfx.dll",
                                L"mso.dll",
                                L"mso20win32client.dll",
                                L"mso30win32client.dll",
                                L"combase.dll",
                                L"ole32.dll",
                                L"oleaut32.dll",
                                L"msvcrt.dll",
                                L"ucrtbase.dll",
                                L"kernel32.dll",
                                L"kernelbase.dll",
                                L"user32.dll",
                                L"gdi32.dll",
                                L"gdi32full.dll"}) {
        world.Add(name);
    }
    return world;
}

/**
 * Parses "4,0x20,8" into a list of offsets.
 *
 * Each entry is an offset at which to read a pointer, following it to the next
 * object. Decimal or 0x-prefixed hex, because a reader working from a
 * disassembler has both in front of them.
 */
std::vector<std::size_t> ParseChain(const std::wstring& text) {
    std::vector<std::size_t> offsets;
    std::wistringstream in(text);
    std::wstring piece;
    while (std::getline(in, piece, L',')) {
        const std::size_t first = piece.find_first_not_of(L" \t");
        if (first == std::wstring::npos) {
            continue;
        }
        const std::size_t last = piece.find_last_not_of(L" \t");
        piece = piece.substr(first, last - first + 1);
        if (piece.empty()) {
            continue;
        }
        const int base =
            (piece.size() > 2 && piece[0] == L'0' && (piece[1] == L'x' || piece[1] == L'X')) ? 16
                                                                                             : 10;
        offsets.push_back(static_cast<std::size_t>(std::wcstoul(piece.c_str(), nullptr, base)));
    }
    return offsets;
}

/**
 * A stable digest of an object's leading words.
 *
 * FNV-1a over the raw bytes. Not cryptographic and does not need to be - it is
 * comparing one object against itself a moment later, and the only property
 * required is that a changed byte changes the result.
 */
std::uint64_t Digest(const void* object, std::size_t bytes) {
    std::uint64_t hash = 1469598103934665603ull;
    const auto* cursor = static_cast<const std::uint8_t*>(object);
    for (std::size_t index = 0; index < bytes; ++index) {
        if (!bb::native::IsReadable(cursor + index, 1)) {
            break;
        }
        hash ^= cursor[index];
        hash *= 1099511628211ull;
    }
    return hash;
}

std::wstring HexBytes(const std::uint8_t* bytes, std::size_t count) {
    std::wostringstream out;
    out << std::hex << std::setfill(L'0');
    for (std::size_t index = 0; index < count; ++index) {
        out << std::setw(2) << static_cast<unsigned>(bytes[index]) << L' ';
    }
    return out.str();
}

/// Dumps one object's leading words, attributing each to a module when it can.
void DumpWords(std::wostringstream& out,
               const bb::native::ModuleWorld& world,
               const void* object,
               std::size_t span) {
    for (std::size_t offset = 0; offset < span; offset += sizeof(void*)) {
        const std::uintptr_t word = bb::native::ReadPointer(object, offset);
        if (word == 0) {
            continue;
        }
        out << L"    +0x" << std::hex << offset << std::dec << L"  0x" << std::hex << word
            << std::dec;
        if (world.Contains(word)) {
            out << L"  " << world.Describe(word);
        } else if (bb::native::IsReadable(reinterpret_cast<const void*>(word), sizeof(void*))) {
            const std::uintptr_t inner =
                bb::native::ReadPointer(reinterpret_cast<const void*>(word), 0);
            if (world.Contains(inner)) {
                out << L"  heap object, vtable " << world.Describe(inner);
            } else {
                out << L"  heap/data";
            }
        }
        out << L"\r\n";
    }
}

} // namespace

std::wstring
inspectChain(IDispatch* fill, const std::wstring& chain, long slotCount, long wordSpan) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing Fill object");
    }
    if (slotCount < 0 || slotCount > 128) {
        slotCount = 24;
    }
    if (wordSpan <= 0 || wordSpan > 0x400) {
        wordSpan = 0x60;
    }

    const bb::native::ModuleWorld world = OfficeWorld();
    const std::vector<std::size_t> offsets = ParseChain(chain);

    std::wostringstream out;
    const void* object = fill;
    out << L"hop 0: Shape.Fill\r\n";

    for (std::size_t hop = 0; hop <= offsets.size(); ++hop) {
        if (!bb::native::IsReadable(object, sizeof(void*))) {
            out << L"  object at hop " << hop << L" is not readable - chain stops here\r\n";
            return out.str();
        }

        const std::uintptr_t vtable = bb::native::ReadPointer(object, 0);
        out << L"  object   0x" << std::hex << reinterpret_cast<std::uintptr_t>(object) << std::dec
            << L"\r\n";
        out << L"  vtable   " << world.Describe(vtable) << L"\r\n";
        out << L"  digest   0x" << std::hex << Digest(object, static_cast<std::size_t>(wordSpan))
            << std::dec << L"\r\n";

        // A vtable that does not validate is the clearest sign this is not the
        // kind of object the walk assumed, and it is worth saying before any of
        // its slots are printed as though they meant something.
        const bb::native::Validation check = bb::native::ValidateVtable(world, vtable, 4);
        out << L"  vtable valid: " << (check.ok ? L"yes" : L"NO - ")
            << (check.ok ? L"" : check.detail) << L"\r\n";

        out << L"  words:\r\n";
        DumpWords(out, world, object, static_cast<std::size_t>(wordSpan));

        if (check.ok && slotCount > 0) {
            out << L"  slots:\r\n";
            for (long slot = 0; slot < slotCount; ++slot) {
                const std::uintptr_t target =
                    bb::native::ReadPointer(reinterpret_cast<const void*>(vtable),
                                            sizeof(void*) * static_cast<std::size_t>(slot));
                if (target == 0) {
                    continue;
                }
                out << L"    [" << std::setw(3) << slot << L"] " << world.Describe(target);
                if (bb::native::IsReadable(reinterpret_cast<const void*>(target), 12)) {
                    out << L"  " << HexBytes(reinterpret_cast<const std::uint8_t*>(target), 12);
                }
                out << L"\r\n";
            }
        }

        if (hop == offsets.size()) {
            break;
        }

        const std::uintptr_t next = bb::native::ReadPointer(object, offsets[hop]);
        out << L"\r\nhop " << (hop + 1) << L": follow +0x" << std::hex << offsets[hop] << std::dec
            << L" -> 0x" << std::hex << next << std::dec << L"\r\n";
        if (next == 0) {
            out << L"  null - chain stops here\r\n";
            return out.str();
        }
        object = reinterpret_cast<const void*>(next);
    }
    return out.str();
}

std::wstring digestChain(IDispatch* fill, const std::wstring& chain, long wordSpan) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing Fill object");
    }
    if (wordSpan <= 0 || wordSpan > 0x400) {
        wordSpan = 0x60;
    }

    const std::vector<std::size_t> offsets = ParseChain(chain);
    const void* object = fill;
    std::wostringstream out;

    /*
     * One line per hop: address, vtable and digest. Small enough that a harness
     * can take it before and after an operation and diff the two, which is how
     * "which object holds the fill" gets answered without calling anything.
     */
    for (std::size_t hop = 0;; ++hop) {
        if (!bb::native::IsReadable(object, sizeof(void*))) {
            out << L"hop" << hop << L"=unreadable;";
            break;
        }
        out << L"hop" << hop << L"=0x" << std::hex << reinterpret_cast<std::uintptr_t>(object)
            << L",vt=0x" << bb::native::ReadPointer(object, 0) << L",digest=0x"
            << Digest(object, static_cast<std::size_t>(wordSpan)) << std::dec << L";";
        if (hop == offsets.size()) {
            break;
        }
        const std::uintptr_t next = bb::native::ReadPointer(object, offsets[hop]);
        if (next == 0) {
            out << L"hop" << (hop + 1) << L"=null;";
            break;
        }
        object = reinterpret_cast<const void*>(next);
    }
    return out.str();
}

namespace {

/// One object found by walking, with enough about it to recognise it again.
struct GraphNode {
    std::uintptr_t address = 0;
    std::uintptr_t vtable = 0;
    std::uint64_t digest = 0;
    std::size_t depth = 0;
    std::uintptr_t parent = 0;
    std::size_t viaOffset = 0;
};

/**
 * Walks outward from an object, collecting everything that looks like an object.
 *
 * Guessing chains by hand does not scale past the second hop, and the question -
 * which object holds the picture fill - is answered far better by walking
 * everything reachable and asking which of it *changed*. So this is a bounded
 * breadth-first walk, and the caller diffs two of them.
 *
 * "Looks like an object" is deliberately strict: the word must point at
 * committed memory whose own first word lands in a known Office module. That
 * rejects strings, floats that happen to look like pointers, and interior
 * pointers, without needing to know anything about what the objects are.
 *
 * Bounded three ways - depth, node count, and a visited set - because the object
 * graph inside PowerPoint is effectively unbounded and a walk that tried to
 * finish would not.
 */
void WalkGraph(const bb::native::ModuleWorld& world,
               const void* root,
               std::size_t maxDepth,
               std::size_t maxNodes,
               std::size_t span,
               std::vector<GraphNode>& nodes) {
    std::vector<std::uintptr_t> visited;
    std::vector<GraphNode> queue;

    const auto seen = [&visited](std::uintptr_t address) {
        for (const std::uintptr_t each : visited) {
            if (each == address) {
                return true;
            }
        }
        return false;
    };

    GraphNode start;
    start.address = reinterpret_cast<std::uintptr_t>(root);
    start.vtable = bb::native::ReadPointer(root, 0);
    start.digest = Digest(root, span);
    queue.push_back(start);
    visited.push_back(start.address);

    for (std::size_t index = 0; index < queue.size() && nodes.size() < maxNodes; ++index) {
        const GraphNode node = queue[index];
        nodes.push_back(node);
        if (node.depth >= maxDepth) {
            continue;
        }

        const auto* object = reinterpret_cast<const void*>(node.address);
        for (std::size_t offset = 0; offset < span; offset += sizeof(void*)) {
            const std::uintptr_t word = bb::native::ReadPointer(object, offset);
            if (word == 0 || seen(word)) {
                continue;
            }
            // A module address is code or static data, not an object to walk.
            if (world.Contains(word)) {
                continue;
            }
            if (!bb::native::IsReadable(reinterpret_cast<const void*>(word), sizeof(void*))) {
                continue;
            }
            const std::uintptr_t innerVtable =
                bb::native::ReadPointer(reinterpret_cast<const void*>(word), 0);
            if (!world.Contains(innerVtable)) {
                continue;
            }

            GraphNode child;
            child.address = word;
            child.vtable = innerVtable;
            child.digest = Digest(reinterpret_cast<const void*>(word), span);
            child.depth = node.depth + 1;
            child.parent = node.address;
            child.viaOffset = offset;
            visited.push_back(word);
            queue.push_back(child);
        }
    }
}

} // namespace

std::wstring digestGraph(IDispatch* fill, long maxDepth, long maxNodes, long wordSpan) {
    if (!fill) {
        throw bb::Error(E_POINTER, "Missing Fill object");
    }
    if (maxDepth <= 0 || maxDepth > 8) {
        maxDepth = 3;
    }
    if (maxNodes <= 0 || maxNodes > 4000) {
        maxNodes = 600;
    }
    if (wordSpan <= 0 || wordSpan > 0x400) {
        wordSpan = 0x80;
    }

    const bb::native::ModuleWorld world = OfficeWorld();
    std::vector<GraphNode> nodes;
    WalkGraph(world,
              fill,
              static_cast<std::size_t>(maxDepth),
              static_cast<std::size_t>(maxNodes),
              static_cast<std::size_t>(wordSpan),
              nodes);

    /*
     * One line per object: address, vtable as module+RVA, digest, and how it was
     * reached. The address is included so a diff can tell a replaced object from
     * a mutated one, and the path is included so a changed node can be found
     * again by hand.
     */
    std::wostringstream out;
    out << L"nodes=" << nodes.size() << L";depth=" << maxDepth << L";span=0x" << std::hex
        << wordSpan << std::dec << L"\r\n";
    for (const GraphNode& node : nodes) {
        out << std::hex << L"0x" << node.address << L'|' << world.Describe(node.vtable) << L"|0x"
            << node.digest << L"|d" << std::dec << node.depth << L"|from 0x" << std::hex
            << node.parent << L"+0x" << node.viaOffset << std::dec << L"\r\n";
    }
    return out.str();
}
