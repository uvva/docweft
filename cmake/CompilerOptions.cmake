# cmake/CompilerOptions.cmake
# Applies compiler warnings and hardening flags per target.
#
# Usage:
#   docweft_apply_compiler_options(<target>)

function(docweft_apply_compiler_options target)
    # ── MSVC (Windows) ───────────────────────────────────────────────────────
    if(MSVC)
        target_compile_options(${target} PRIVATE
            /W4             # High warning level
            /WX             # Warnings as errors
            /utf-8          # Source and execution charset = UTF-8
            /permissive-    # Strict standards conformance
            /Zc:__cplusplus # Report correct __cplusplus value
            /wd4251         # class member needs dll-interface — benign for STL across same MSVC toolchain
            /wd4275         # non dll-interface base (e.g. std::runtime_error) — same rationale as 4251
        )
        target_compile_definitions(${target} PRIVATE
            _CRT_SECURE_NO_WARNINGS     # silence strncpy/sprintf nags
            NOMINMAX                    # don't let windows.h define min/max macros
            WIN32_LEAN_AND_MEAN         # trim windows.h
            UNICODE
            _UNICODE
        )

    # ── Clang / GCC (Linux, macOS) ───────────────────────────────────────────
    else()
        target_compile_options(${target} PRIVATE
            -Wall
            -Wextra
            -Wpedantic
            -Wshadow
            -Wconversion
            -Wsign-conversion
            -Wnull-dereference
            -Wdouble-promotion
            -Wformat=2
            # Configuration-specific
            $<$<CONFIG:Debug>:  -g -O0 -fno-omit-frame-pointer>
            $<$<CONFIG:Release>:-O3 -DNDEBUG>
            $<$<CONFIG:RelWithDebInfo>:-O2 -g -DNDEBUG>
        )

        # Clang-specific extras
        if(CMAKE_CXX_COMPILER_ID MATCHES "Clang")
            target_compile_options(${target} PRIVATE
                -Wno-c++98-compat         # we target C++20, no need for 98 compat noise
                -Wno-c++98-compat-pedantic
            )
        endif()

        # macOS: suppress deprecation of BSD APIs where clang is overly noisy
        if(APPLE)
            target_compile_definitions(${target} PRIVATE
                _DARWIN_C_SOURCE
            )
        endif()
    endif()

    # ── Export header (works for both STATIC and SHARED) ─────────────────────
    # When STATIC: DOCWEFT_STATIC_DEFINE is added → DOCWEFT_API is empty.
    # When SHARED: DOCWEFT_API expands to __declspec(dllexport/dllimport)
    #              on Windows, __attribute__((visibility("default"))) elsewhere.
    include(GenerateExportHeader)
    generate_export_header(${target}
        EXPORT_MACRO_NAME     DOCWEFT_API
        EXPORT_FILE_NAME      ${CMAKE_CURRENT_BINARY_DIR}/export/docweft/export.h
        DEPRECATED_MACRO_NAME DOCWEFT_DEPRECATED
    )
    target_include_directories(${target} PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}/export>
        $<INSTALL_INTERFACE:include>
    )

    # Signal STATIC-builds to the macro machinery.
    get_target_property(_tf_type ${target} TYPE)
    if(_tf_type STREQUAL "STATIC_LIBRARY")
        target_compile_definitions(${target} PUBLIC DOCWEFT_STATIC_DEFINE)
    endif()

    # ── Default-hidden visibility for SHARED builds (Linux/macOS) ────────────
    # Only explicitly DOCWEFT_API-annotated symbols get exported.
    # Matches the dllexport-on-demand behavior of Windows.
    set_target_properties(${target} PROPERTIES
        CXX_VISIBILITY_PRESET    hidden
        C_VISIBILITY_PRESET      hidden
        VISIBILITY_INLINES_HIDDEN ON
        POSITION_INDEPENDENT_CODE ON   # PIC needed when linking static deps into a .so
    )
endfunction()
