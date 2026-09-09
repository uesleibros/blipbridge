/**
 * @file receiver_inspect.cpp
 * Reports the OART receiver reachable from a PowerPoint FillFormat.
 *
 * The walk itself, and every guard around it, lives in oart_layout.cpp so this
 * file and native_apply.cpp cannot drift apart. This one only formats what the
 * walk found, for comparing Shapes against each other.
 *
 * Nothing here calls a private Office function. Every pointer reported is
 * borrowed and belongs to Office; no reference is taken and nothing is retained
 * after the call returns. In particular a receiver must never be cached: a Shape
 * deleted through public COM still passes every check in the chain.
 */

#include "../experiment_api.hpp"
#include "../../src/backend/windows_office/oart_layout.hpp"

#include <blipbridge/dispatch.hpp>

#include <cstring>
#include <sstream>

namespace {

/// The receiver fields worth comparing between Shapes; see docs/receiver_lookup.md.
constexpr std::size_t kReceiverContainerOffset = 0x08;
constexpr std::size_t kReceiverSequenceOffset = 0x20;
/// The OART FillFormat's non-atomic COM reference count, which makes this STA-only.
constexpr std::size_t kFillFormatReferenceCountOffset = 0x30;

void AppendPointer(std::wostringstream& out, const wchar_t* label, std::uintptr_t value) {
    out << label << L"=0x" << std::hex << value << std::dec << L';';
}

} // namespace

std::wstring inspectFillReceiver(IDispatch* fill) {
    // Report the environment before walking, so a failure downstream still tells
    // the reader which build they are on and what the walk expected to find.
    std::wostringstream environment;
    environment << L"expectedBuild=" << bb::oart::kSupportedVersionText
                << L";oartVersion=" << bb::oart::ModuleVersionText(L"oart.dll")
                << L";ppcoreVersion=" << bb::oart::ModuleVersionText(L"ppcore.dll")
                << L";gfxVersion=" << bb::oart::ModuleVersionText(L"gfx.dll")
                << L";expectedOartFillFormatVtable=oart.dll+0xAF60B8"
                << L";expectedOartReceiverVtable=oart.dll+0x9F6658"
                << L";observedPpcoreFillFormatVtable=ppcore.dll+0x1464478"
                << L";ppcoreVtableIsStructural=1;";

    bb::oart::FillTarget target;
    try {
        target = bb::oart::ResolveFillTarget(fill);
    } catch (const bb::Error& error) {
        // Re-throw with the environment attached: the guard message alone does
        // not say which build drifted or what the walk expected.
        const std::wstring wide = environment.str();
        throw bb::Error(error.hr, std::string(error.what()) + " | " +
                                      std::string(wide.begin(), wide.end()));
    }

    std::wostringstream out;
    out << environment.str();
    out << L"build=" << bb::oart::kSupportedVersionText << L';';
    out << L"ppcoreFillFormatVtable=ppcore.dll+0x" << std::hex << target.wrapper.vtableRva
        << std::dec << L";wrapperInnerOffset=0x" << std::hex << target.wrapper.innerOffset
        << std::dec << L";wrapperThunks=" << target.wrapper.identityThunks << L';';
    AppendPointer(out, L"oart", target.oartBase);
    AppendPointer(out, L"publicFill", reinterpret_cast<std::uintptr_t>(target.publicFill));
    AppendPointer(out, L"handler", reinterpret_cast<std::uintptr_t>(target.handler));
    out << L"handlerRefs="
        << bb::oart::LoadDword(target.handler, kFillFormatReferenceCountOffset) << L';';
    AppendPointer(out, L"token", reinterpret_cast<std::uintptr_t>(target.token));
    out << L"tokenStrong=" << target.tokenStrong << L';';
    AppendPointer(out, L"receiver", reinterpret_cast<std::uintptr_t>(target.receiver));
    AppendPointer(out, L"container",
                  bb::oart::LoadPointer(target.receiver, kReceiverContainerOffset));
    out << L"sequence=" << bb::oart::LoadPointer(target.receiver, kReceiverSequenceOffset)
        << L';';
    return out.str();
}
