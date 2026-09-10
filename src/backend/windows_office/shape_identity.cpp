/**
 * @file shape_identity.cpp
 * The one implementation of Shape identity for caching.
 *
 * Moved here unchanged from picture_cache.cpp when a second cache needed the
 * same rules. The group exclusion below is the part that matters most: it is
 * there because filling a group provably changes what its children render, and
 * the two caches must never disagree about that.
 */

#include "shape_identity.hpp"

#include <blipbridge/dispatch.hpp>
#include <blipbridge/errors.hpp>

#include <cstdint>
#include <string>

namespace bb::office {

ShapeKey DescribeShape(IDispatch* shape, long shapeType) noexcept {
    ShapeKey key;
    // A group's fill propagates to its children, and a child's fill changes what
    // the group shows, so a remembered texture on either side goes stale when the
    // other is filled. Neither is cached; see the file comment for the pixel
    // evidence. msoGroup is 6.
    constexpr long kGroup = 6;
    if (shapeType == kGroup) {
        return key;
    }
    try {
        // ParentGroup answers only for a Shape inside a group. One property read,
        // about a microsecond, against the 190 an incorrect skip would misplace.
        bb::Value owner = bb::get(shape, L"ParentGroup");
        if (owner.v.vt == VT_DISPATCH && owner.obj()) {
            return key;
        }
    } catch (const bb::Error&) {
        // Top-level Shapes refuse the question, which is the common case and
        // means exactly what it should: this Shape is not inside a group.
    }
    try {
        key.shape = bb::get(shape, L"Id").integer();
        bb::Value parent = bb::get(shape, L"Parent");
        if (parent.v.vt != VT_DISPATCH || !parent.obj()) {
            return key;
        }
        // Shape.Parent is the Slide. Anything that does not report a SlideID
        // goes uncached rather than being given a key that means something else.
        key.slide = bb::get(parent.obj(), L"SlideID").integer();
        bb::Value presentation = bb::get(parent.obj(), L"Parent");
        if (presentation.v.vt != VT_DISPATCH || !presentation.obj()) {
            return key;
        }
        // Presentations have no numeric id, but every open one has a distinct
        // window-independent hash of its full name plus its index. The index
        // alone would shift as documents open and close.
        bb::Value name = bb::get(presentation.obj(), L"FullName");
        std::wstring text;
        if (name.v.vt == VT_BSTR && name.v.bstrVal) {
            text.assign(name.v.bstrVal, SysStringLen(name.v.bstrVal));
        }
        if (text.empty()) {
            // An unsaved presentation has no FullName. Its Name ("Presentation1")
            // is unique among open documents, which is all this key needs.
            bb::Value shortName = bb::get(presentation.obj(), L"Name");
            if (shortName.v.vt == VT_BSTR && shortName.v.bstrVal) {
                text.assign(shortName.v.bstrVal, SysStringLen(shortName.v.bstrVal));
            }
        }
        if (text.empty()) {
            return key;
        }
        // FNV-1a over the document's name. Unsigned throughout: the constants
        // do not fit a 32-bit long, and signed overflow would be undefined.
        std::uint32_t hash = 2166136261u;
        for (wchar_t character : text) {
            hash ^= static_cast<std::uint32_t>(character);
            hash *= 16777619u;
        }
        key.presentation = static_cast<long>(hash);
        key.valid = true;
    } catch (...) {
        key.valid = false;
    }
    return key;
}

} // namespace bb::office
