# cmake/Dependencies.cmake
# Resolves all third-party dependencies.
#
# Strategy (in order):
#   1. find_package() — satisfied by vcpkg toolchain (Windows/macOS)
#      or Nix / system packages (Linux).
#   2. FetchContent fallback — only when DOCWEFT_USE_FETCHCONTENT=ON
#      and the package was not found via find_package.
#
# Required packages
#   - pugixml          XML parsing   (vcpkg: pugixml    | nixpkgs: pugixml)
#   - libzip           ZIP archives  (vcpkg: libzip     | nixpkgs: libzip)
#   - nlohmann_json    JSON          (vcpkg: nlohmann-json | nixpkgs: nlohmann_json)
#   - fmt              Formatting    (vcpkg: fmt        | nixpkgs: fmt)
#   - inja             Templates     (vcpkg: inja       | nixpkgs: — use FetchContent)
#
# Optional (auto-detected, feature-flagged into the library):
#   - stb_image /      JPEG + BMP decode, PNG encode. Single-header, public
#     stb_image_write  domain. Pulled via FetchContent from nothings/stb when
#                      not found through find_package. If missing, setImage
#                      falls back to PNG-only (NotImplemented for JPG/BMP).
#                      Adds DOCWEFT_HAVE_STB compile definition.
#   - TIFF (libtiff)   TIFF decode (vcpkg: tiff | nixpkgs: libtiff). If not
#                      present, setImage throws NotImplemented for TIFF.
#                      Adds DOCWEFT_HAVE_TIFF compile definition.
#   - PoDoFo           Native OOXML→PDF rendering with no Microsoft Word or
#     (opt-in via       LibreOffice installed — see PLAN.md. Only searched
#     DOCWEFT_        when DOCWEFT_ENABLE_NATIVE_PDF=ON (default OFF);
#     ENABLE_NATIVE_PDF) fails the configure step loudly if enabled but not
#                        found, rather than silently degrading, since a user
#                        who opted in clearly wants this path to work.
#                        (vcpkg: podofo | nixpkgs: podofo). Build it with its
#                        MPL-2.0 licensing option selected, not the
#                        LGPL-2.0-or-later alternative — keeps static linking
#                        free of relinking obligations (see PLAN.md).
#                        Adds DOCWEFT_HAVE_PODOFO compile definition.
#
# Optional (tests only):
#   - Catch2 3.x       (vcpkg: catch2 | nixpkgs: catch2_3)

include(FetchContent)

# When building docweft as SHARED on Linux/macOS, all static deps we pull
# into it must be compiled with -fPIC. Setting this globally before
# FetchContent_MakeAvailable() ensures pugixml/libzip/fmt/etc. pick it up.
if(BUILD_SHARED_LIBS OR DOCWEFT_BUILD_SHARED)
    set(CMAKE_POSITION_INDEPENDENT_CODE ON)
endif()

# ── Helper: print dep status ────────────────────────────────────────────────
macro(_tf_dep_status name found_var)
    if(${found_var})
        message(STATUS "  [dep] ${name}: found via find_package")
    else()
        message(STATUS "  [dep] ${name}: NOT found — will use FetchContent")
    endif()
endmacro()

# ── pugixml ─────────────────────────────────────────────────────────────────
find_package(pugixml CONFIG QUIET)
_tf_dep_status("pugixml" pugixml_FOUND)

if(NOT pugixml_FOUND)
    if(DOCWEFT_USE_FETCHCONTENT)
        FetchContent_Declare(pugixml
            GIT_REPOSITORY https://github.com/zeux/pugixml.git
            GIT_TAG        v1.14
            GIT_SHALLOW    TRUE
        )
        FetchContent_MakeAvailable(pugixml)
    else()
        message(WARNING "pugixml not found — stubs will compile without it. Install via vcpkg/Nix or set DOCWEFT_USE_FETCHCONTENT=ON")
    endif()
endif()

# ── libzip ───────────────────────────────────────────────────────────────────
# vcpkg exports as "libzip::zip"; pkg-config may use "libzip"
find_package(libzip CONFIG QUIET)
if(NOT libzip_FOUND)
    find_package(PkgConfig QUIET)
    if(PkgConfig_FOUND)
        pkg_check_modules(LIBZIP IMPORTED_TARGET libzip)
        if(LIBZIP_FOUND)
            # Provide a uniform alias so CMakeLists.txt doesn't need ifdefs
            add_library(libzip::zip ALIAS PkgConfig::LIBZIP)
            set(libzip_FOUND TRUE)
        endif()
    endif()
endif()
_tf_dep_status("libzip" libzip_FOUND)

if(NOT libzip_FOUND)
    if(DOCWEFT_USE_FETCHCONTENT)
        FetchContent_Declare(libzip
            GIT_REPOSITORY https://github.com/nih-at/libzip.git
            GIT_TAG        v1.10.1
            GIT_SHALLOW    TRUE
        )
        set(ENABLE_COMMONCRYPTO OFF CACHE BOOL "" FORCE)
        set(ENABLE_GNUTLS       OFF CACHE BOOL "" FORCE)
        set(ENABLE_MBEDTLS      OFF CACHE BOOL "" FORCE)
        set(ENABLE_OPENSSL      ON  CACHE BOOL "" FORCE)
        set(BUILD_TOOLS         OFF CACHE BOOL "" FORCE)
        set(BUILD_REGRESS       OFF CACHE BOOL "" FORCE)
        set(BUILD_EXAMPLES      OFF CACHE BOOL "" FORCE)
        set(BUILD_DOC           OFF CACHE BOOL "" FORCE)
        FetchContent_MakeAvailable(libzip)
    else()
        message(WARNING "libzip not found — stubs will compile without it.")
    endif()
endif()

# ── nlohmann_json ────────────────────────────────────────────────────────────
# Version-gated: the namespaced nlohmann_json::nlohmann_json target we link
# against below only exists from 3.2.0 on — older packages only expose the
# bare, un-namespaced target.
find_package(nlohmann_json 3.2.0 CONFIG QUIET)
_tf_dep_status("nlohmann_json" nlohmann_json_FOUND)

if(NOT nlohmann_json_FOUND)
    if(DOCWEFT_USE_FETCHCONTENT)
        FetchContent_Declare(nlohmann_json
            GIT_REPOSITORY https://github.com/nlohmann/json.git
            GIT_TAG        v3.11.3
            GIT_SHALLOW    TRUE
        )
        set(JSON_BuildTests OFF CACHE BOOL "" FORCE)
        FetchContent_MakeAvailable(nlohmann_json)
    else()
        message(WARNING "nlohmann_json not found — stubs will compile without it.")
    endif()
endif()

# ── fmt ──────────────────────────────────────────────────────────────────────
find_package(fmt CONFIG QUIET)
_tf_dep_status("fmt" fmt_FOUND)

if(NOT fmt_FOUND)
    if(DOCWEFT_USE_FETCHCONTENT)
        FetchContent_Declare(fmt
            GIT_REPOSITORY https://github.com/fmtlib/fmt.git
            GIT_TAG        10.2.1
            GIT_SHALLOW    TRUE
        )
        FetchContent_MakeAvailable(fmt)
    else()
        message(WARNING "fmt not found — stubs will compile without it.")
    endif()
endif()

# ── inja ─────────────────────────────────────────────────────────────────────
# inja is in vcpkg; not packaged in nixpkgs → always try FetchContent as fallback.
find_package(inja CONFIG QUIET)
_tf_dep_status("inja" inja_FOUND)

if(NOT inja_FOUND)
    if(DOCWEFT_USE_FETCHCONTENT)
        FetchContent_Declare(inja
            GIT_REPOSITORY https://github.com/pantor/inja.git
            GIT_TAG        v3.4.0
            GIT_SHALLOW    TRUE
        )
        set(INJA_BUILD_TESTS      OFF CACHE BOOL "" FORCE)
        set(INJA_EXPORT           OFF CACHE BOOL "" FORCE)
        set(BUILD_BENCHMARK       OFF CACHE BOOL "" FORCE)

        # inja's own CMakeLists.txt (checked at build/_deps/inja-src, v3.4.0):
        # with INJA_USE_EMBEDDED_JSON=OFF, it does `if(TARGET nlohmann_json)
        # install(TARGETS nlohmann_json ...)` — it assumes any target
        # literally named `nlohmann_json` (no namespace) was defined locally
        # via add_library (e.g. a git submodule) and is therefore
        # installable. That holds when *we* FetchContent'd nlohmann_json
        # above (its own CMakeLists.txt defines that exact target the same
        # way), but not when find_package() resolved it to a package
        # config's target instead (nixpkgs/vcpkg/...): several distros'
        # nlohmann_jsonConfig.cmake also registers a bare `nlohmann_json`
        # IMPORTED target for pre-3.2.0-style compatibility, and
        # `install(TARGETS)` on an IMPORTED target is a hard CMake error —
        # "given target ... which does not exist" — regardless of version.
        #
        # So only force the switch off in the FetchContent case. When
        # find_package() already succeeded, leave inja's own default (ON)
        # alone: its ON-branch calls find_package(nlohmann_json) itself,
        # resolves to the very same package we already found, and never
        # reaches the install(TARGETS) line at all.
        if(NOT nlohmann_json_FOUND)
            set(INJA_USE_EMBEDDED_JSON OFF CACHE BOOL "" FORCE)  # use our FetchContent'd nlohmann_json, not inja's bundled copy
        endif()
        FetchContent_MakeAvailable(inja)
    else()
        message(WARNING "inja not found — stubs will compile without it.")
    endif()
endif()

# ── stb (Stage 5 Phase 2: JPEG + BMP decode, PNG encode) ────────────────────
# Single-header public domain library from https://github.com/nothings/stb.
# vcpkg exposes it as `Stb::Stb` via find_package(Stb CONFIG); nix has `stb`
# but without CMake config — so we FetchContent as a fallback.
#
# Offline/sandboxed builds (no network access for FetchContent's git clone):
# point FETCHCONTENT_SOURCE_DIR_STB at a local checkout of the stb sources
# and FetchContent_Populate below uses it directly instead of cloning. This
# is a stock CMake FetchContent mechanism (works for any FetchContent_Declare
# name), not something DocWeft implements itself.
set(_tf_stb_include "")
find_package(Stb QUIET)
if(Stb_FOUND AND DEFINED Stb_INCLUDE_DIR)
    set(_tf_stb_include "${Stb_INCLUDE_DIR}")
    message(STATUS "  [dep] stb: found via find_package (${Stb_INCLUDE_DIR})")
elseif(DOCWEFT_USE_FETCHCONTENT)
    message(STATUS "  [dep] stb: NOT found — fetching nothings/stb")
    # Pinned, not "master": nothings/stb has no release tags, and stb_image.h's
    # JPEG decoder is exactly the kind of unmanaged-memory C parser where an
    # unpinned "always latest" dependency means nobody can say which version
    # actually shipped when a crash is reported (see
    # DOCWEFT_UPSTREAM_FEEDBACK.md item 10). Every other FetchContent dep
    # above is pinned to an exact tag; this was the one exception.
    # Commit below is nothings/stb@master as of 2026-09-11 — bump deliberately,
    # not by drifting.
    FetchContent_Declare(stb
        GIT_REPOSITORY https://github.com/nothings/stb.git
        GIT_TAG        2c980bb59875b0d32144a71867fbdebb2f77cd20
        GIT_SHALLOW    TRUE
    )
    FetchContent_GetProperties(stb)
    if(NOT stb_POPULATED)
        FetchContent_Populate(stb)
    endif()
    if(EXISTS "${stb_SOURCE_DIR}/stb_image.h")
        set(_tf_stb_include "${stb_SOURCE_DIR}")
    else()
        message(WARNING "  [dep] stb: FetchContent completed but stb_image.h not found — JPEG/BMP support disabled")
    endif()
else()
    message(WARNING "  [dep] stb: NOT found and FetchContent disabled — JPEG/BMP support disabled; setImage() will throw NotImplemented for non-PNG input")
endif()

if(_tf_stb_include)
    # Expose as plain cache vars rather than an INTERFACE library, so the
    # install(EXPORT ...) set doesn't need to include an intermediate target
    # that points at a build-tree-only include dir.
    set(DOCWEFT_STB_INCLUDE_DIR "${_tf_stb_include}" CACHE INTERNAL "")
    set(DOCWEFT_HAVE_STB        ON                   CACHE INTERNAL "")
else()
    set(DOCWEFT_STB_INCLUDE_DIR "" CACHE INTERNAL "")
    set(DOCWEFT_HAVE_STB        OFF CACHE INTERNAL "")
endif()

# ── libtiff (Stage 5 Phase 2: TIFF decode) ──────────────────────────────────
find_package(TIFF QUIET)
if(TIFF_FOUND)
    message(STATUS "  [dep] TIFF: found via find_package")
    set(DOCWEFT_HAVE_TIFF ON  CACHE INTERNAL "TIFF decoder available")
else()
    message(STATUS "  [dep] TIFF: NOT found — TIFF support disabled")
    set(DOCWEFT_HAVE_TIFF OFF CACHE INTERNAL "TIFF decoder available")
endif()

# ── PoDoFo (Stage 6b: native PDF renderer, opt-in) ──────────────────────────
if(DOCWEFT_ENABLE_NATIVE_PDF)
    find_package(podofo CONFIG QUIET)
    if(podofo_FOUND)
        message(STATUS "  [dep] PoDoFo: found via find_package")
        set(DOCWEFT_HAVE_PODOFO ON CACHE INTERNAL "Native PDF renderer available")
    else()
        message(FATAL_ERROR
            "DOCWEFT_ENABLE_NATIVE_PDF=ON but PoDoFo was not found "
            "(vcpkg: podofo | nixpkgs: podofo). Install it or turn the "
            "option off.")
    endif()
else()
    set(DOCWEFT_HAVE_PODOFO OFF CACHE INTERNAL "Native PDF renderer available")
endif()

# ── Catch2 (tests only) ──────────────────────────────────────────────────────
if(DOCWEFT_BUILD_TESTS)
    find_package(Catch2 3 CONFIG QUIET)
    _tf_dep_status("Catch2 v3" Catch2_FOUND)

    if(NOT Catch2_FOUND)
        if(DOCWEFT_USE_FETCHCONTENT)
            FetchContent_Declare(Catch2
                GIT_REPOSITORY https://github.com/catchorg/Catch2.git
                GIT_TAG        v3.5.3
                GIT_SHALLOW    TRUE
            )
            FetchContent_MakeAvailable(Catch2)
            list(APPEND CMAKE_MODULE_PATH "${catch2_SOURCE_DIR}/extras")
        else()
            message(FATAL_ERROR "Catch2 v3 not found. Set DOCWEFT_USE_FETCHCONTENT=ON to auto-download.")
        endif()
    endif()
endif()
