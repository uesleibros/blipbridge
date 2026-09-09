cmake_minimum_required(VERSION 3.20)
project(BlipBridge VERSION 0.1.0 LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
if(NOT WIN32 OR NOT CMAKE_SIZEOF_VOID_P EQUAL 8)
  message(FATAL_ERROR "Initial research build requires Windows x64")
endif()
add_compile_definitions(UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
add_compile_options(-Wall -Wextra -Wpedantic)
include_directories(include)
add_executable(bb_experiment experiments/main.cpp)
target_link_libraries(bb_experiment PRIVATE ole32 oleaut32 uuid psapi)
target_link_options(bb_experiment PRIVATE -municode -static-libgcc -static-libstdc++)
add_library(BlipBridge SHARED
    # Public C ABI and the platform seam beneath it.
    src/abi/blipbridge_abi.cpp
    src/backend/windows_office_backend.cpp
    src/dllmain.cpp
    src/com/class_factory.cpp
    src/com/engine.cpp
    src/com/engine_textures.cpp
    src/image/validate.cpp
    # Explicit research entry points; never selected by the normal texture API.
    experiments/automation_bridge.cpp
    experiments/main.cpp
    experiments/exp_userpicture/trace.cpp
    experiments/exp_pickup_apply/focused.cpp
    experiments/exp_pickup_apply/stress.cpp
    experiments/exp_internal_blip/receiver_inspect.cpp
    experiments/exp_internal_blip/cached_image_load.cpp
    experiments/exp_internal_blip/oart_layout.cpp
    experiments/exp_internal_blip/native_apply.cpp
    experiments/exp_internal_blip/native_texture.cpp
    experiments/exp_internal_blip/texture_benchmark.cpp
    experiments/exp_internal_blip/batch_benchmark.cpp
    experiments/exp_internal_blip/pixel_texture.cpp
)
target_compile_definitions(BlipBridge PRIVATE BB_DLL BLIPBRIDGE_BUILD)
target_link_libraries(BlipBridge PRIVATE ole32 oleaut32 uuid psapi version windowscodecs)
target_link_options(BlipBridge PRIVATE -static -static-libgcc -static-libstdc++)
set_target_properties(BlipBridge PROPERTIES PREFIX "")
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

include(CTest)
if(BUILD_TESTING)
    add_executable(bb_com_contract tests/com_contract.cpp)
    target_link_libraries(bb_com_contract PRIVATE ole32 oleaut32 uuid)
    target_link_options(bb_com_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
    add_test(NAME com_contract COMMAND bb_com_contract $<TARGET_FILE:BlipBridge>)

    # The C ABI is exercised the way the VBA wrapper binds to it: LoadLibraryW
    # by path, then GetProcAddress by undecorated name. No COM involved.
    add_executable(bb_abi_contract tests/abi_contract.cpp)
    target_link_options(bb_abi_contract PRIVATE -municode -static -static-libgcc -static-libstdc++)
    add_test(NAME abi_contract COMMAND bb_abi_contract $<TARGET_FILE:BlipBridge>)
endif()
