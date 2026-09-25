cmake_minimum_required(VERSION 3.21)

project(DocWeft
    VERSION 0.1.0
    DESCRIPTION "Cross-platform library for DOCX template manipulation"
    HOMEPAGE_URL "https://github.com/uvva/docweft"
    LANGUAGES CXX
)

# ============================================================
# Global settings
# ============================================================
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# Default build type
if(NOT CMAKE_BUILD_TYPE AND NOT CMAKE_CONFIGURATION_TYPES)
    set(CMAKE_BUILD_TYPE Release CACHE STRING "Build type" FORCE)
    set_property(CACHE CMAKE_BUILD_TYPE PROPERTY STRINGS Debug Release RelWithDebInfo MinSizeRel)
endif()

# ============================================================
# Options
# ============================================================
option(DOCWEFT_BUILD_TESTS    "Build the test suite"               ON)
option(DOCWEFT_BUILD_EXAMPLES "Build example programs"             OFF)
option(DOCWEFT_BUILD_SHARED   "Build as shared library"            OFF)
option(DOCWEFT_USE_FETCHCONTENT "Fall back to FetchContent for deps not found via find_package" ON)
option(DOCWEFT_ENABLE_NATIVE_PDF "Native OOXML->PDF rendering via PoDoFo, no Microsoft Word/LibreOffice required (see PLAN.md)" OFF)
option(DOCWEFT_ENABLE_REMOTE_CONVERTER "HTTP converter fallback (Gotenberg-compatible endpoint from DOCWEFT_CONVERTER_URL, via the curl CLI)" OFF)

# ============================================================
# Platform diagnostics
# ============================================================
if(WIN32)
    message(STATUS "[DocWeft] Platform: Windows — use vcpkg toolchain")
    message(STATUS "  cmake -B build -DCMAKE_TOOLCHAIN_FILE=<vcpkg>/scripts/buildsystems/vcpkg.cmake")
elseif(APPLE)
    message(STATUS "[DocWeft] Platform: macOS — use vcpkg or Homebrew")
elseif(UNIX)
    message(STATUS "[DocWeft] Platform: Linux — use Nix dev shell (flake.nix) or system packages")
endif()

# ============================================================
# CMake modules
# ============================================================
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

include(cmake/Dependencies.cmake)
include(cmake/CompilerOptions.cmake)

# ============================================================
# Library target
# ============================================================
if(DOCWEFT_BUILD_SHARED)
    add_library(docweft SHARED)
else()
    add_library(docweft STATIC)
endif()

add_library(DocWeft::docweft ALIAS docweft)

# stb_image / libtiff are third-party C decoders that have crashed (not
# thrown) on malformed/corrupt input in the field — a hardware exception
# (e.g. access violation), not a catchable ReportException. /EHa switches
# image.cpp to the asynchronous exception model so _set_se_translator()
# (see src/docx/image.cpp) can actually intercept that and rethrow it as a
# normal C++ exception instead of letting it kill the process. Scoped to
# this one file; nothing else needs the (slower) async model. No-op on
# non-MSVC compilers, which don't recognize /EHa as a flag they own here.
if(MSVC)
    set_source_files_properties(src/docx/image.cpp PROPERTIES COMPILE_OPTIONS /EHa)
endif()

target_sources(docweft
    PRIVATE
        src/docweft.cpp
        src/error.cpp
        src/docx/merger.cpp
        src/docx/image.cpp
        src/docx/pdf_native.cpp
        src/docx/reader.cpp
        src/docx/writer.cpp
        src/template/engine.cpp
)

target_include_directories(docweft
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# v1 requires all core libs. Verify each target exists after Dependencies.cmake
# has run; fail loudly if not (FetchContent fallback is enabled by default).
set(_tf_required_deps
    pugixml::pugixml              # XML parsing (DOCX internals)
    libzip::zip                   # ZIP archive handling (DOCX container)
    nlohmann_json::nlohmann_json  # JSON data for templates
    fmt::fmt                      # string formatting
)
# Optional deps linked only when available. Each one flips a HAVE_* compile
# definition so image.cpp can conditionally include decoder code.
# (stb is handled separately below — it's headers-only and exposed as a plain
# include dir, not a CMake target, to keep the install export set clean.)
set(_tf_optional_deps
    inja::inja            # Stage 3+: Jinja2-style placeholders
    TIFF::TIFF            # Stage 5 Phase 2: TIFF decode
    podofo::podofo         # Stage 6b: native PDF rendering (opt-in, see PLAN.md)
)

foreach(_dep IN LISTS _tf_required_deps)
    if(NOT TARGET ${_dep})
        message(FATAL_ERROR "Required dependency target '${_dep}' not found. "
                            "Enable DOCWEFT_USE_FETCHCONTENT=ON or install via vcpkg/Nix.")
    endif()
endforeach()

set(_tf_optional_resolved "")
foreach(_dep IN LISTS _tf_optional_deps)
    if(TARGET ${_dep})
        list(APPEND _tf_optional_resolved ${_dep})
    endif()
endforeach()

target_link_libraries(docweft
    PUBLIC
        # nothing public-facing yet — keep the ABI clean
    PRIVATE
        ${_tf_required_deps}
        ${_tf_optional_resolved}
)

# Feature-flag the optional decoders. stb is included via BUILD_INTERFACE so
# it never leaks into the installed consumer's include paths; libtiff is linked
# normally (its CMake target properly propagates through install).
if(DOCWEFT_HAVE_STB)
    target_include_directories(docweft
        PRIVATE $<BUILD_INTERFACE:${DOCWEFT_STB_INCLUDE_DIR}>)
    target_compile_definitions(docweft PRIVATE DOCWEFT_HAVE_STB=1)
endif()
if(TARGET TIFF::TIFF)
    target_compile_definitions(docweft PRIVATE DOCWEFT_HAVE_TIFF=1)
endif()
if(TARGET podofo::podofo)
    target_compile_definitions(docweft PRIVATE DOCWEFT_HAVE_PODOFO=1)
endif()
if(DOCWEFT_ENABLE_REMOTE_CONVERTER)
    target_compile_definitions(docweft PRIVATE DOCWEFT_HAVE_REMOTE_CONVERTER=1)
endif()

docweft_apply_compiler_options(docweft)

set_target_properties(docweft PROPERTIES
    VERSION   ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}
)

# ============================================================
# Internal static twin (for tests that need access to internal symbols).
# Only built when building SHARED — otherwise the main target is already STATIC.
# ============================================================
if(DOCWEFT_BUILD_SHARED AND DOCWEFT_BUILD_TESTS)
    add_library(docweft_internal STATIC
        src/docweft.cpp
        src/error.cpp
        src/docx/merger.cpp
        src/docx/image.cpp
        src/docx/pdf_native.cpp
        src/docx/reader.cpp
        src/docx/writer.cpp
        src/template/engine.cpp
    )
    add_library(DocWeft::docweft_internal ALIAS docweft_internal)

    target_include_directories(docweft_internal
        PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}/export>
        PRIVATE
            ${CMAKE_CURRENT_SOURCE_DIR}/src
    )
    # STATIC ⇒ macro expansions are empty ⇒ all symbols are visible.
    target_compile_definitions(docweft_internal PUBLIC DOCWEFT_STATIC_DEFINE)

    target_link_libraries(docweft_internal
        PUBLIC
            ${_tf_required_deps}
            ${_tf_optional_resolved}
    )
    # Same HAVE_* propagation as the main target. PUBLIC here so the test TU
    # that links docweft_internal also sees these flags — the BMP/TIFF
    # test cases need DOCWEFT_HAVE_STB at compile time.
    if(DOCWEFT_HAVE_STB)
        target_include_directories(docweft_internal
            PUBLIC $<BUILD_INTERFACE:${DOCWEFT_STB_INCLUDE_DIR}>)
        target_compile_definitions(docweft_internal PUBLIC DOCWEFT_HAVE_STB=1)
    endif()
    if(TARGET TIFF::TIFF)
        target_compile_definitions(docweft_internal PUBLIC DOCWEFT_HAVE_TIFF=1)
    endif()
    if(TARGET podofo::podofo)
        target_compile_definitions(docweft_internal PUBLIC DOCWEFT_HAVE_PODOFO=1)
    endif()
    if(DOCWEFT_ENABLE_REMOTE_CONVERTER)
        target_compile_definitions(docweft_internal PUBLIC DOCWEFT_HAVE_REMOTE_CONVERTER=1)
    endif()

    set_target_properties(docweft_internal PROPERTIES
        POSITION_INDEPENDENT_CODE ON
    )
endif()

# ============================================================
# Tests
# ============================================================
if(DOCWEFT_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()

# ============================================================
# Examples
# ============================================================
if(DOCWEFT_BUILD_EXAMPLES)
    add_subdirectory(examples)
endif()

# ============================================================
# Install / packaging
# ============================================================
include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

# On a Windows SHARED build `docweft.dll` links privately against fmt.dll,
# pugixml.dll, tiff.dll, zlib1.dll, libzip.dll, etc. — `cmake --install` needs
# to copy those alongside our own DLL so the installed prefix is self-contained.
# RUNTIME_DEPENDENCIES (CMake 3.21+) scans the binary for imports and resolves
# them through DIRECTORIES; the exclude regexes keep Windows system DLLs out.
# On Linux/macOS we keep this off: CMake's default POST_EXCLUDE filters system
# paths, but for our Nix path the transitive .so set is provided by the
# consumer's environment, not by our prefix.
#
# Under vcpkg (VCPKG_TOOLCHAIN is set by its toolchain file) this must stay
# off entirely, not just be redirected: vcpkg already installs fmt/pugixml/
# libzip/tiff's DLLs into the shared installed/<triplet>/(debug/)bin pool
# that consumers read from directly. Pointing DIRECTORIES at that pool makes
# RUNTIME_DEPENDENCIES copy second copies of those DLLs into our own install
# tree, which vcpkg's post-build validation then rejects as "already
# installed by <port>" ownership conflicts. There's no "resolve but don't
# copy" flag, so not entering this block is the only correct fix.
set(_tf_runtime_deps "")
if(WIN32 AND DOCWEFT_BUILD_SHARED AND NOT VCPKG_TOOLCHAIN)
    set(_tf_runtime_deps
        RUNTIME_DEPENDENCIES
            PRE_EXCLUDE_REGEXES  "api-ms-.*" "ext-ms-.*"
            POST_EXCLUDE_REGEXES ".*[Ss]ystem32.*" ".*SYSWOW64.*"
            DIRECTORIES          "$<TARGET_FILE_DIR:docweft>"
    )
endif()

install(TARGETS docweft
    EXPORT   DocWeftTargets
    ${_tf_runtime_deps}
    ARCHIVE  DESTINATION ${CMAKE_INSTALL_LIBDIR} COMPONENT Development
    LIBRARY  DESTINATION ${CMAKE_INSTALL_LIBDIR} COMPONENT Runtime
             NAMELINK_COMPONENT                  Development
    RUNTIME  DESTINATION ${CMAKE_INSTALL_BINDIR} COMPONENT Runtime
    INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

# Public headers authored in the source tree.
install(DIRECTORY include/
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
    COMPONENT   Development
)

# Auto-generated export header. GenerateExportHeader writes it into the build
# dir only; without this install step every consumer gets a compile error on
# `#include "docweft/export.h"` (used by merger.hpp / error.hpp /
# docweft.hpp for the DOCWEFT_API dllexport macro).
install(FILES "${CMAKE_CURRENT_BINARY_DIR}/export/docweft/export.h"
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/docweft
    COMPONENT   Development
)

install(EXPORT DocWeftTargets
    FILE        DocWeftTargets.cmake
    NAMESPACE   DocWeft::
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/DocWeft
    COMPONENT   Development
)

configure_package_config_file(
    cmake/DocWeftConfig.cmake.in
    ${CMAKE_CURRENT_BINARY_DIR}/DocWeftConfig.cmake
    INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/DocWeft
)

write_basic_package_version_file(
    ${CMAKE_CURRENT_BINARY_DIR}/DocWeftConfigVersion.cmake
    VERSION       ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)

install(FILES
    ${CMAKE_CURRENT_BINARY_DIR}/DocWeftConfig.cmake
    ${CMAKE_CURRENT_BINARY_DIR}/DocWeftConfigVersion.cmake
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/DocWeft
    COMPONENT   Development
)

# Build-tree export: projects pulling DocWeft via add_subdirectory() or
# FetchContent can call find_package(DocWeft) against the build dir without
# installing first.
export(EXPORT DocWeftTargets
    FILE      "${CMAKE_CURRENT_BINARY_DIR}/DocWeftTargets.cmake"
    NAMESPACE DocWeft::
)
