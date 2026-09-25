{
  description = "DocWeft — cross-platform C++20 DOCX template library";

  inputs = {
    nixpkgs.url     = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, flake-utils }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = nixpkgs.legacyPackages.${system};

        # Packages available in nixpkgs for all the deps we need.
        # Note: inja is NOT in nixpkgs → CMake will fall back to FetchContent.
        buildDeps = with pkgs; [
          pugixml        # XML parser — nixpkgs attr: "pugixml"
          libzip         # ZIP archive — nixpkgs attr: "libzip"
          nlohmann_json  # JSON        — nixpkgs attr: "nlohmann_json"
          fmt            # Formatting  — nixpkgs attr: "fmt"
          libtiff        # TIFF decode (Stage 5) — nixpkgs attr: "libtiff"
          # PoDoFo: native OOXML->PDF rendering with no MS Word/LibreOffice
          # installed, opt-in via -DDOCWEFT_ENABLE_NATIVE_PDF=ON (see
          # PLAN.md). Dual-licensed LGPL-2.0-or-later OR MPL-2.0 upstream —
          # this project consumes it under the MPL-2.0 option to keep static
          # linking free of relinking obligations.
          podofo         # nixpkgs attr: "podofo"
          # inja: not packaged → fetched by CMake FetchContent automatically.
          # stb: not packaged as a CMake config — fetched via FetchContent.
        ];

        testDeps = with pkgs; [
          catch2_3       # Catch2 v3   — nixpkgs attr: "catch2_3"
        ];

        nativeBuildDeps = with pkgs; [
          cmake
          ninja
          pkg-config
        ];

        devTools = with pkgs; [
          # Unversioned: follows nixpkgs' default LLVM, so the shell doesn't
          # break when an old pinned version (e.g. clang_17) is dropped.
          clang
          clang-tools      # clang-tidy, clang-format, clangd
          gdb
          valgrind
          cmake-format
          python3          # for scripts / cmake --trace parsing
        ];

      in {
        # ─── Development shell ────────────────────────────────────────────────
        # Enter with: nix develop
        devShells.default = pkgs.mkShell {
          name = "docweft-dev";

          packages = nativeBuildDeps ++ buildDeps ++ testDeps ++ devTools;

          shellHook = ''
            echo ""
            echo "  DocWeft dev environment (Linux / Nix)"
            echo ""
            echo "  Build (Release):"
            echo "    cmake -B build -G Ninja -DCMAKE_BUILD_TYPE=Release"
            echo "    cmake --build build -j$(nproc)"
            echo ""
            echo "  Build (Debug + tests):"
            echo "    cmake -B build-debug -G Ninja -DCMAKE_BUILD_TYPE=Debug"
            echo "    cmake --build build-debug && ctest --test-dir build-debug -V"
            echo ""
            echo "  Note: 'inja' is not in nixpkgs — CMake will fetch it via FetchContent."
            echo "        Set DOCWEFT_USE_FETCHCONTENT=OFF to disable this fallback."
            echo ""
          '';
        };

        # ─── Library derivation ───────────────────────────────────────────────
        # Build with: nix build
        packages.default = pkgs.stdenv.mkDerivation {
          pname   = "docweft";
          version = "0.1.0";
          src     = ./.;

          nativeBuildInputs = nativeBuildDeps;
          buildInputs       = buildDeps;

          cmakeFlags = [
            "-DCMAKE_BUILD_TYPE=Release"
            "-DDOCWEFT_BUILD_TESTS=OFF"
            "-DDOCWEFT_BUILD_EXAMPLES=OFF"
            # FetchContent is allowed so inja can be fetched during the build.
            # In a hermetic Nix build you may want to pre-vendor inja instead.
            "-DDOCWEFT_USE_FETCHCONTENT=ON"
          ];

          meta = with pkgs.lib; {
            description = "Cross-platform C++20 library for DOCX template manipulation";
            license     = licenses.mit;
            platforms   = platforms.unix;
          };
        };

        # ─── Formatter / linter check ─────────────────────────────────────────
        # Run with: nix run .#format-check
        apps.format-check = flake-utils.lib.mkApp {
          drv = pkgs.writeShellScriptBin "format-check" ''
            set -euo pipefail
            echo "Checking C++ formatting..."
            find src include tests -name '*.cpp' -o -name '*.hpp' | \
              xargs ${pkgs.clang-tools}/bin/clang-format --dry-run --Werror
            echo "All files formatted correctly."
          '';
        };
      }
    );
}
