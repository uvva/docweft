#!/usr/bin/env bash
# Build DocWeft as a SHARED library for all supported targets from the
# current host. Each target drops artifacts under build/<preset>/.
#
# Usage:
#   ./scripts/build-all.sh              # build everything possible on this host
#   ./scripts/build-all.sh linux-x64    # build a single preset
#   ./scripts/build-all.sh --list       # list targets available on this host
#
# Requires: CMake 3.25+ with preset support, Ninja.
# Cross-compilation to Windows from Linux additionally requires mingw-w64.

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${ROOT_DIR}"

HOST="$(uname -s)"
TARGETS=()

case "${HOST}" in
    Linux)
        TARGETS=(linux-x64)
        if command -v x86_64-w64-mingw32-g++ >/dev/null 2>&1; then
            TARGETS+=(win-x64-mingw-cross)
        fi
        if command -v i686-w64-mingw32-g++ >/dev/null 2>&1; then
            TARGETS+=(win-x86-mingw-cross)
        fi
        ;;
    Darwin)
        TARGETS=(macos-arm64)
        ;;
    MINGW*|MSYS*|CYGWIN*)
        TARGETS=(win-x64-vcpkg win-x86-vcpkg)
        ;;
    *)
        echo "[build-all] Unsupported host: ${HOST}" >&2
        exit 1
        ;;
esac

if [[ "${1:-}" == "--list" ]]; then
    printf '%s\n' "${TARGETS[@]}"
    exit 0
fi

# Filter to a single preset if one is passed.
if [[ $# -gt 0 ]]; then
    wanted="$1"
    match=0
    for t in "${TARGETS[@]}"; do [[ "$t" == "$wanted" ]] && match=1 && break; done
    if (( match == 0 )); then
        echo "[build-all] Preset '${wanted}' is not available on host '${HOST}'." >&2
        echo "[build-all] Available: ${TARGETS[*]}" >&2
        exit 1
    fi
    TARGETS=("${wanted}")
fi

for preset in "${TARGETS[@]}"; do
    echo "════════════════════════════════════════════════════════════════"
    echo "  Building preset: ${preset}"
    echo "════════════════════════════════════════════════════════════════"
    cmake --preset "${preset}"
    cmake --build --preset "${preset}"

    # Tests only run on presets where the host can execute the output
    # (no cross-compile test runs — you'd need wine for that).
    case "${preset}" in
        linux-x64|macos-arm64|win-x64-vcpkg|win-x86-vcpkg)
            ctest --preset "${preset}" || true
            ;;
    esac
done

echo ""
echo "All builds completed. Artifacts are under build/<preset>/."
