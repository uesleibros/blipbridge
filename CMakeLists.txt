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
    single build produces. The two architectures ship different sets of files:

      x64  the validated PowerPoint backend, the COM research surface, and the
           research probes that produced them
      x86  the public C ABI over a backend that refuses, and nothing else

    That exclusion is the safety property. The 64-bit backend depends on the x64
    calling convention, on per-build module RVAs and on object layouts 32-bit
    Office does not share, so compiling it as x86 would produce something that
    links and is wrong - the single most dangerous outcome for a library that
    drives undocumented Office internals. It is not compiled at all.
]]
if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(BB_ARCH "x64")
elseif(CMAKE_SIZEOF_VOID_P EQUAL 4)
    set(BB_ARCH "x86")
else()
    message(FATAL_ERROR "Unrecognised pointer size ${CMAKE_SIZEOF_VOID_P}")
endif()

# True only where a validated native Office backend exists for this architecture.
if(BB_ARCH STREQUAL "x64")
    set(BB_NATIVE_BACKEND ON)
else()
    set(BB_NATIVE_BACKEND OFF)
endif()

message(STATUS "BlipBridge version           : ${PROJECT_VERSION}")
message(STATUS "BlipBridge architecture      : ${BB_ARCH}")
if(BB_NATIVE_BACKEND)
    message(STATUS "BlipBridge native backend    : validated PowerPoint x64 backend")
else()
    message(STATUS "BlipBridge native backend    : none - the library will load and "
                   "refuse every texture call with a specific reason")
    message(STATUS "BlipBridge x86 status        : see docs/windows_x86.md")
endif()

add_compile_definitions(UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
add_compile_options(-Wall -Wextra -Wpedantic)
include_directories(include)

# --- the library -------------------------------------------------------------
#
# Portable core: the public C ABI, the platform seam, and image validation.
# Nothing here knows what Office is.
set(BB_SOURCES
    src/abi/blipbridge_abi.cpp
    src/image/validate.cpp
    # Software resampling: ordinary portable image maths, no Office, so it is
    # built and tested on every architecture.
    src/image/resample.cpp
    src/dllmain.cpp
)

if(BB_NATIVE_BACKEND)
    list(APPEND BB_SOURCES
        # The seam's Windows implementation.
        src/backend/windows_office_backend.cpp
        # The reverse-engineered backend, reached only through backend.hpp.
        src/backend/windows_office/oart_layout.cpp
        src/backend/windows_office/native_apply.cpp
        src/backend/windows_office/native_texture.cpp
        src/backend/windows_office/shape_policy.cpp
        src/backend/windows_office/picture_cache.cpp
        # The COM Automation surface: research and compatibility, never the
        # normal path. It drives the same texture store the C ABI drives, which
        # is what lets the Office harnesses assert shipping behaviour.
        src/com/class_factory.cpp
        src/com/engine.cpp
        src/com/engine_textures.cpp
        # Explicit research entry points, reachable only through that surface.
        experiments/automation_bridge.cpp
        experiments/main.cpp
        experiments/exp_userpicture/trace.cpp
        experiments/exp_pickup_apply/focused.cpp
        experiments/exp_pickup_apply/stress.cpp
        experiments/exp_internal_blip/receiver_inspect.cpp
        experiments/exp_internal_blip/native_apply_probe.cpp
        experiments/exp_internal_blip/cached_image_load.cpp
        experiments/exp_internal_blip/texture_benchmark.cpp
        experiments/exp_internal_blip/batch_benchmark.cpp
        experiments/exp_internal_blip/pixel_texture.cpp
        experiments/exp_internal_blip/stage_profiler.cpp
        experiments/exp_internal_blip/shape_compatibility.cpp
        experiments/exp_internal_blip/picture_cache_driver.cpp
    )
else()
    list(APPEND BB_SOURCES src/backend/unsupported_backend.cpp)
endif()

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
    # The portable core needs only the OLE allocator and string types.
    target_link_libraries(BlipBridge PRIVATE ole32 oleaut32)
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

    # A timing run, not an assertion: built so it cannot rot, but never added as
    # a test, because a busy runner must not turn a measurement into a failure.
    add_executable(bb_resample_benchmark tests/resample_benchmark.cpp src/image/resample.cpp)
    target_link_options(bb_resample_benchmark PRIVATE -static -static-libgcc -static-libstdc++)

    add_executable(bb_abi_contract tests/abi_contract.cpp)
    target_link_options(bb_abi_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
    add_test(NAME abi_contract COMMAND bb_abi_contract $<TARGET_FILE:BlipBridge>)

    # The COM contract asks the Automation surface for its DISPIDs, so it only
    # exists where that surface is compiled in.
    if(BB_NATIVE_BACKEND)
        add_executable(bb_com_contract tests/com_contract.cpp)
        target_link_libraries(bb_com_contract PRIVATE ole32 oleaut32 uuid)
        target_link_options(bb_com_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
        add_test(NAME com_contract COMMAND bb_com_contract $<TARGET_FILE:BlipBridge>)
    endif()
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
