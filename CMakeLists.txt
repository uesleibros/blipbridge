cmake_minimum_required(VERSION 3.20)

# The public header is the single source of the release number. Parsing it keeps
# the build, the DLL's own BB_GetVersion and the packaged archive names in step
# with one edit rather than three that quietly disagreed.
file(READ "${CMAKE_CURRENT_SOURCE_DIR}/include/blipbridge/blipbridge.h" BB_HEADER)
string(REGEX MATCH "BB_VERSION_MAJOR ([0-9]+)" _ "${BB_HEADER}")
set(BB_VERSION_MAJOR ${CMAKE_MATCH_1})
string(REGEX MATCH "BB_VERSION_MINOR ([0-9]+)" _ "${BB_HEADER}")
set(BB_VERSION_MINOR ${CMAKE_MATCH_1})
string(REGEX MATCH "BB_VERSION_PATCH ([0-9]+)" _ "${BB_HEADER}")
set(BB_VERSION_PATCH ${CMAKE_MATCH_1})

project(BlipBridge
        VERSION ${BB_VERSION_MAJOR}.${BB_VERSION_MINOR}.${BB_VERSION_PATCH}
        LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT WIN32)
    message(FATAL_ERROR
        "BlipBridge builds on Windows only. See docs/macos.md for what a macOS "
        "backend would require - it is research, not a port.")
endif()

#[[
    Architecture is decided by the toolchain, and is then explicit everywhere.

    It is deliberately *not* a source-level `#ifdef` that quietly changes what a
    single build produces. The two architectures ship different backends:

      x64  the accelerated backend - reverse-engineered OART and GFX internals,
           plus the COM research surface and the probes that produced them
      x86  the portable backend - the same public API over documented Office
           Automation, with no Office internals compiled into it at all

    That exclusion is the safety property, and it is why there are two backends
    rather than one with an `#ifdef` in it. The accelerated backend depends on
    the x64 calling convention, on per-build module RVAs and on object layouts
    32-bit Office does not share, so compiling it as x86 would produce something
    that links and is wrong - the single most dangerous outcome for a library
    that drives undocumented Office internals. It is not compiled there at all.

    Three source sets, and the distinction between the first two is the whole
    point:

      BB_OFFICE_PORTABLE_SOURCES  documented Automation. Architecture-neutral,
                                  built everywhere Office is.
      BB_OFFICE_NATIVE_SOURCES    reverse-engineered internals. x64 only.
      BB_BACKEND_SOURCES          whichever backend this architecture gets.
]]
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(BB_ARCH "x64")
elseif(CMAKE_SIZEOF_VOID_P EQUAL 4)
    set(BB_ARCH "x86")
else()
    message(FATAL_ERROR "Unrecognised pointer size ${CMAKE_SIZEOF_VOID_P}")
endif()

# True only where a validated accelerated Office backend exists for this
# architecture. Where it is off the portable backend is built instead, so the
# library is fully functional either way - the difference is speed, not features.
if(BB_ARCH STREQUAL "x64")
    set(BB_NATIVE_BACKEND ON)
else()
    set(BB_NATIVE_BACKEND OFF)
endif()

#[[
    Builds the portable backend on x64 as well, replacing the accelerated one.

    Off by default and never set for a release: this exists so the portable
    backend can be exercised against a real PowerPoint. The only 32-bit-specific
    thing about that backend is the compiler, so running it in the 64-bit host
    tests every line of it that could be wrong, which is a far stronger statement
    than "it compiles for x86".

    A build with this on is a test build. It is asserted against in CI so it can
    never be what ships.
]]
option(BB_FORCE_PORTABLE_BACKEND
       "Build the portable Automation backend even where the accelerated one exists" OFF)
if(BB_FORCE_PORTABLE_BACKEND)
    set(BB_NATIVE_BACKEND OFF)
    message(STATUS "BlipBridge backend override   : portable backend forced - this is a test build")
endif()

message(STATUS "BlipBridge version           : ${PROJECT_VERSION}")
message(STATUS "BlipBridge architecture      : ${BB_ARCH}")
if(BB_NATIVE_BACKEND)
    message(STATUS "BlipBridge backend           : accelerated - validated PowerPoint x64 internals")
else()
    message(STATUS "BlipBridge backend           : portable - documented Office Automation")
    message(STATUS "BlipBridge x86 status        : see docs/windows_x86.md")
endif()

add_compile_definitions(UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
add_compile_options(-Wall -Wextra -Wpedantic)
include_directories(include)

# --- the shipped library -----------------------------------------------------
#
# BB_SOURCES is *only* what ships. The COM Automation surface and the
# reverse-engineering probes are deliberately absent: they belong to the research
# target below, which is a separate DLL that never leaves this repository.
#
# Keeping them apart matters for more than tidiness. The shipped DLL is what runs
# inside someone's PowerPoint, and it should contain nothing that was written to
# poke at Office internals for investigation.
set(BB_SOURCES
    src/abi/blipbridge_abi.cpp
    src/image/validate.cpp
    # Software resampling: ordinary portable image maths, no Office, so it is
    # built and tested on every architecture.
    src/image/resample.cpp
    # Encoded-image decoding through WIC: still ordinary image work with no
    # Office in it, so it builds and is testable on both architectures.
    src/image/decode.cpp
    # Crop, orient, resize: composed here so one decode serves all three.
    src/image/pipeline.cpp
    # The CPU image resource: decoded pixels a caller can process repeatedly,
    # kept deliberately separate from the Office texture resource.
    src/image/store.cpp
    # Projective quad mapping, viable because Office's own picture-fill
    # mapping was measured: see docs and tools/probe_freeform_uv_mapping.ps1.
    src/image/warp.cpp
    src/dllmain.cpp
)

#[[
    The native compatibility framework: architecture detection, module identity,
    guarded reads, structural validation, profiles and the resolver.

    Built on every architecture, and that is the point of it. It is the frame a
    native backend plugs its own facts into, and it decides only whether a
    profile may be used - never what one says. Probing 32-bit Office showed why
    that separation is load-bearing: x64's central structural invariant, a vtable
    of identity thunks, does not exist there at all, so a frame that knew about
    thunks would have been a frame only one architecture could use.

    Nothing here calls an Office internal. It reads, validates, and refuses.
]]
set(BB_NATIVE_COMMON_SOURCES
    src/backend/native/common/architecture.cpp
    src/backend/native/common/memory_safety.cpp
    src/backend/native/common/module_identity.cpp
    src/backend/native/common/structural_validation.cpp
    src/backend/native/common/profile.cpp
    src/backend/native/common/resolver.cpp
)

#[[
    Documented Office Automation, shared by both backends.

    Nothing here reads an Office internal. These three decide which Shape classes
    may be filled, which Shapes may be cached, and when an apply may be skipped -
    and both backends ask the *same* code, so the two architectures cannot
    disagree about a Shape. A second implementation for x86 would have been a
    second set of answers, and the difference would have surfaced to a caller as
    "it behaves differently on 32-bit" with nothing to point at.
]]
set(BB_OFFICE_PORTABLE_SOURCES
    src/backend/windows_office/shape_policy.cpp
    src/backend/windows_office/shape_identity.cpp
    src/backend/windows_office/apply_skip_cache.cpp
)

# The reverse-engineered internals. x64 only - see the note at the top.
set(BB_OFFICE_NATIVE_SOURCES
    src/backend/windows_office_backend.cpp
    src/backend/windows_office/oart_layout.cpp
    src/backend/windows_office/native_apply.cpp
    src/backend/windows_office/native_texture.cpp
    src/backend/windows_office/range_texture.cpp
    src/backend/windows_office/picture_cache.cpp
)

# The portable backend: the same public API over Fill.UserPicture.
set(BB_OFFICE_PORTABLE_BACKEND_SOURCES
    src/backend/portable_office_backend.cpp
    src/backend/portable_office/portable_texture.cpp
    # PNG encoding exists only for this backend: Fill.UserPicture takes a path,
    # so a texture built from raw pixels has to become a file.
    src/image/encode.cpp
)

#[[
    The COM Automation surface and the probes it exposes. Research only: this is
    how the Office harnesses drive the same library the C ABI drives, which is
    what lets them assert shipping behaviour without shipping themselves.

    It is split the same way the backends are, and for the same reason. The
    portable set drives the public C ABI and nothing else, so it builds over
    either backend - which is what lets the Office harnesses run against the
    portable backend in a real PowerPoint. The native set instruments Office
    internals: receiver layouts, transaction splitting, GFX image lifetimes.
    There is nothing for those to instrument without the accelerated backend, so
    they are not built there, and their Automation ids answer with that reason.
]]
set(BB_RESEARCH_SOURCES
    src/com/class_factory.cpp
    src/com/engine.cpp
    experiments/automation_bridge.cpp
    experiments/main.cpp
    experiments/exp_userpicture/trace.cpp
    experiments/exp_pickup_apply/focused.cpp
    experiments/exp_pickup_apply/stress.cpp
    experiments/exp_internal_blip/batch_benchmark.cpp
    experiments/exp_internal_blip/skip_benchmark.cpp
    experiments/exp_internal_blip/apply_cost_factors.cpp
    experiments/exp_internal_blip/range_apply.cpp
    experiments/exp_internal_blip/image_api_driver.cpp
    experiments/exp_internal_blip/picture_cache_driver.cpp
    # Read-only reconnaissance for the 32-bit native backend. Architecture-
    # neutral on purpose: the same source runs on both, so a difference in its
    # output is a difference in Office rather than in the instrument.
    experiments/exp_native_x86/structure_probe.cpp
)

set(BB_RESEARCH_NATIVE_SOURCES
    src/com/engine_textures.cpp
    experiments/exp_internal_blip/receiver_inspect.cpp
    experiments/exp_internal_blip/native_apply_probe.cpp
    experiments/exp_internal_blip/cached_image_load.cpp
    experiments/exp_internal_blip/texture_benchmark.cpp
    experiments/exp_internal_blip/resolve_profiler.cpp
    experiments/exp_internal_blip/transaction_split.cpp
    experiments/exp_internal_blip/quad_apply.cpp
    experiments/exp_internal_blip/dynamic_probe.cpp
    experiments/exp_internal_blip/pixel_texture.cpp
    experiments/exp_internal_blip/stage_profiler.cpp
    experiments/exp_internal_blip/shape_compatibility.cpp
    experiments/exp_internal_blip/apply_profiler.cpp
)

if(BB_NATIVE_BACKEND)
    list(APPEND BB_RESEARCH_SOURCES ${BB_RESEARCH_NATIVE_SOURCES})
else()
    list(APPEND BB_RESEARCH_SOURCES src/com/engine_textures_portable.cpp)
endif()

#[[
    Exactly one backend is compiled, and the choice is made here rather than
    inside the sources. WIN32 is required above, so the only question left is
    which Windows backend - and both of them fill Shapes. The difference is
    whether Office internals or Office's own API does it.

    src/backend/unsupported_backend.cpp is therefore not built by any
    configuration today. It is kept because it is the shape a macOS backend
    would start from and docs/macos.md points at it; the fatal error above is
    what actually enforces the platform, not that file.
]]
list(APPEND BB_SOURCES ${BB_OFFICE_PORTABLE_SOURCES})
list(APPEND BB_SOURCES ${BB_NATIVE_COMMON_SOURCES})
if(BB_NATIVE_BACKEND)
    set(BB_BACKEND_SOURCES ${BB_OFFICE_NATIVE_SOURCES})
else()
    set(BB_BACKEND_SOURCES ${BB_OFFICE_PORTABLE_BACKEND_SOURCES})
endif()
list(APPEND BB_SOURCES ${BB_BACKEND_SOURCES})

add_library(BlipBridge SHARED ${BB_SOURCES})
target_compile_definitions(BlipBridge PRIVATE BB_DLL BLIPBRIDGE_BUILD)
target_link_options(BlipBridge PRIVATE -static -static-libgcc -static-libstdc++)
#[[
    The DLL carries its architecture in its name, so both can sit beside the same
    presentation and BlipBridge.bas can load the one matching the PowerPoint
    process. Naming it here rather than only at packaging time means a local
    build is the same file the wrapper looks for.
]]
set_target_properties(BlipBridge PROPERTIES
    PREFIX ""
    OUTPUT_NAME "BlipBridge-${BB_ARCH}")
if(BB_ARCH STREQUAL "x86")
    #[[
        x86 exports are stdcall, because VBA's Declare can call nothing else.
        --kill-at strips the @N suffix stdcall would otherwise add, so the
        exported names match x64's exactly and one VBA Declare works on both.
        Without it every export would be `_BB_Init@0` and no Declare would bind.
    ]]
    target_link_options(BlipBridge PRIVATE -Wl,--kill-at)
endif()
if(BB_NATIVE_BACKEND)
    target_link_libraries(BlipBridge PRIVATE ole32 oleaut32 uuid psapi version windowscodecs)
else()
    #[[
        The portable core needs the OLE allocator, string types, and WIC - the
        decoder is image work rather than Office work, so it exists on x86 too.

        psapi and version are for the native compatibility framework, which is
        built on every architecture: it reads module identities in order to
        decide that no native profile applies. Refusing to accelerate is a
        decision that has to be made on both, so its dependencies are needed on
        both.
    ]]
    target_link_libraries(BlipBridge PRIVATE ole32 oleaut32 psapi version windowscodecs)
endif()

#[[
    The research library.

    Never packaged, never released. It compiles the same production sources plus
    the COM surface, so the Office harnesses exercise the shipping code paths
    through an Automation object - and the DLL a user installs stays free of them.

    It is built for whichever backend this configuration has, which is the whole
    reason the portable backend can be validated at all: with
    BB_FORCE_PORTABLE_BACKEND the harnesses drive the portable backend inside a
    real PowerPoint, rather than only proving that it compiles.
]]
add_library(BlipBridgeResearch SHARED ${BB_SOURCES} ${BB_RESEARCH_SOURCES})
target_compile_definitions(BlipBridgeResearch PRIVATE BB_DLL BLIPBRIDGE_BUILD)
if(BB_NATIVE_BACKEND)
    # Gates the probes that read Office internals. The Automation ids stay put
    # either way - they are a stable ABI - but without this they answer with the
    # reason they are missing instead of running.
    target_compile_definitions(BlipBridgeResearch PRIVATE BB_HAS_NATIVE_BACKEND)
endif()
target_link_options(BlipBridgeResearch PRIVATE -static -static-libgcc -static-libstdc++)
set_target_properties(BlipBridgeResearch PROPERTIES
    PREFIX ""
    OUTPUT_NAME "BlipBridgeResearch-${BB_ARCH}")
target_link_libraries(BlipBridgeResearch
                      PRIVATE ole32 oleaut32 uuid psapi version windowscodecs)
if(BB_ARCH STREQUAL "x86")
    #[[
        The same reason the shipped DLL needs it, for a different caller.
        DllGetClassObject and DllCanUnloadNow are stdcall, so on x86 they would
        export as `DllGetClassObject@12` - and COM looks them up undecorated. A
        research DLL without this loads and then cannot be registered as a COM
        server at all, which is how it would be discovered: not at build time,
        but by a harness that could never reach it.
    ]]
    target_link_options(BlipBridgeResearch PRIVATE -Wl,--kill-at)
endif()

# --- tests -------------------------------------------------------------------
include(CTest)
if(BUILD_TESTING)
    #[[
        The ABI contract runs on any architecture and needs no Office. On x86 it
        is the whole test surface, and it is exactly the right one: it asserts
        that the library loads, exports every entry point undecorated, and fails
        closed with a readable reason - which is precisely what an x86 build is
        supposed to do today.
    ]]
    #[[
        The resampling filters are pure image maths, so this runs everywhere and
        is the only suite that can catch a filter regression - nothing
        downstream would notice a wrong pixel.
    ]]
    add_executable(bb_resample_contract tests/resample_contract.cpp src/image/resample.cpp)
    target_link_options(bb_resample_contract PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME resample_contract COMMAND bb_resample_contract)

    # Encoded-image decoding. Needs WIC but not Office, so it runs on both
    # architectures - the first piece of BlipBridge that genuinely does.
    add_executable(bb_decode_contract tests/decode_contract.cpp
                   src/image/decode.cpp src/image/resample.cpp)
    target_link_libraries(bb_decode_contract PRIVATE ole32 oleaut32 windowscodecs)
    target_link_options(bb_decode_contract PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME decode_contract COMMAND bb_decode_contract)

    #[[
        The native compatibility framework, and above all its refusals.

        Runs off-host on both architectures: the framework knows nothing about
        Office, so its policy can be driven with a fake architecture resolver and
        asserted exactly. Almost every check here is a way it must decline.
    ]]
    add_executable(bb_native_framework_contract tests/native_framework_contract.cpp
                   ${BB_NATIVE_COMMON_SOURCES})
    target_link_libraries(bb_native_framework_contract PRIVATE psapi version)
    target_link_options(bb_native_framework_contract PRIVATE
                        -static -static-libgcc -static-libstdc++)
    add_test(NAME native_framework_contract COMMAND bb_native_framework_contract)

    #[[
        PNG encoding, and the round trip through it.

        Encoding exists only for the portable backend, and the portable backend
        is what makes 32-bit PowerPoint work - so this runs on both
        architectures, and on x86 it is testing the code that actually ships
        there.
    ]]
    add_executable(bb_encode_contract tests/encode_contract.cpp
                   src/image/encode.cpp src/image/decode.cpp src/image/resample.cpp)
    target_link_libraries(bb_encode_contract PRIVATE ole32 oleaut32 windowscodecs)
    target_link_options(bb_encode_contract PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME encode_contract COMMAND bb_encode_contract)

    # Crop, orient and resize, against numbered-grid fixtures so every assertion
    # says which source pixel arrived where.
    add_executable(bb_pipeline_contract tests/pipeline_contract.cpp
                   src/image/pipeline.cpp src/image/decode.cpp src/image/resample.cpp)
    target_link_libraries(bb_pipeline_contract PRIVATE ole32 oleaut32 windowscodecs)
    target_link_options(bb_pipeline_contract PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME pipeline_contract COMMAND bb_pipeline_contract)

    # The projective warp. Pure maths, so it needs neither Office nor WIC.
    add_executable(bb_warp_contract tests/warp_contract.cpp
                   src/image/warp.cpp src/image/resample.cpp)
    target_link_options(bb_warp_contract PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME warp_contract COMMAND bb_warp_contract)

    # A timing run, not an assertion: built so it cannot rot, never registered
    # as a test, because a busy runner must not fail a measurement.
    add_executable(bb_image_benchmark tests/image_benchmark.cpp src/image/pipeline.cpp
                   src/image/decode.cpp src/image/resample.cpp src/image/warp.cpp)
    target_link_libraries(bb_image_benchmark PRIVATE ole32 oleaut32 windowscodecs)
    target_link_options(bb_image_benchmark PRIVATE -static -static-libgcc -static-libstdc++)

    # A timing run, not an assertion: built so it cannot rot, but never added as
    # a test, because a busy runner must not turn a measurement into a failure.
    add_executable(bb_resample_benchmark tests/resample_benchmark.cpp src/image/resample.cpp)
    target_link_options(bb_resample_benchmark PRIVATE -static -static-libgcc -static-libstdc++)

    #[[
        The 64-bit handle's journey across the VBA boundary, reproduced in C++.
        Runs on both architectures and needs no Office - and no VBA, which is
        exactly its limit: it proves the arithmetic and the slot layout, not that
        a real 32-bit PowerPoint executes the wrapper correctly.
    ]]
    add_executable(bb_handle_marshalling tests/handle_marshalling.cpp)
    target_link_options(bb_handle_marshalling PRIVATE -static -static-libgcc -static-libstdc++)
    add_test(NAME handle_marshalling COMMAND bb_handle_marshalling)

    add_executable(bb_abi_contract tests/abi_contract.cpp)
    target_link_options(bb_abi_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
    add_test(NAME abi_contract COMMAND bb_abi_contract $<TARGET_FILE:BlipBridge>)

    #[[
        The COM contract asks the Automation surface for its DISPIDs. That
        surface is now built for whichever backend this configuration has, so
        this runs everywhere too - and on x86 it is the one test that exercises
        the dispatch table in a 32-bit process, where the ids are the same
        numbers but the marshalling is not.
    ]]
    add_executable(bb_com_contract tests/com_contract.cpp)
    target_link_libraries(bb_com_contract PRIVATE ole32 oleaut32 uuid)
    target_link_options(bb_com_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
    add_test(NAME com_contract COMMAND bb_com_contract $<TARGET_FILE:BlipBridgeResearch>)
endif()

# --- research tools ----------------------------------------------------------
#
# Reverse-engineering aids, not shipped. They exist only where there is a
# validated backend to investigate.
if(BB_NATIVE_BACKEND)
    add_executable(bb_experiment experiments/main.cpp)
    target_link_libraries(bb_experiment PRIVATE ole32 oleaut32 uuid psapi)
    target_link_options(bb_experiment PRIVATE -municode -static-libgcc -static-libstdc++)

    add_executable(bb_symbols tools/symbolize.cpp)
    target_link_libraries(bb_symbols PRIVATE dbghelp)
    target_link_options(bb_symbols PRIVATE -municode -static -static-libgcc -static-libstdc++)

    add_executable(bb_typelib tools/typelib.cpp experiments/main.cpp)
    target_compile_definitions(bb_typelib PRIVATE BB_DLL)
    target_link_libraries(bb_typelib PRIVATE ole32 oleaut32 uuid psapi)
    target_link_options(bb_typelib PRIVATE -municode -static -static-libgcc -static-libstdc++)

    add_executable(bb_interfaces tools/interfaces.cpp)
    target_link_libraries(bb_interfaces PRIVATE ole32 oleaut32 uuid)
    target_link_options(bb_interfaces PRIVATE -municode -static -static-libgcc -static-libstdc++)

    add_executable(bb_peanchors tools/pe_anchors.cpp)
    target_link_options(bb_peanchors PRIVATE -municode -static -static-libgcc -static-libstdc++)
endif()
