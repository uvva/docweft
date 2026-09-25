#pragma once

#include <cstdint>
#include <filesystem>
#include <string>

namespace docweft::docx {

/// Image format detected by magic bytes.
enum class ImageFormat {
    Unknown,
    Png,
    Jpeg,   // Stage 5 Phase 2 — decoded via stb_image, re-encoded as PNG
    Bmp,    // Stage 5 Phase 2 — decoded via stb_image, re-encoded as PNG
    Tiff,   // Stage 5 Phase 2 — decoded via libtiff, re-encoded as PNG
};

/// Raw PNG bytes plus dimensions in pixels.
struct PngBuffer {
    std::string bytes;
    std::uint32_t width  = 0;
    std::uint32_t height = 0;
};

/// Sniff the first bytes of `path` and return a matching ImageFormat.
/// Returns Unknown if the file does not exist or doesn't match any known format.
[[nodiscard]] ImageFormat detect_format(const std::filesystem::path& path);

/// Load an image file, normalizing it to a PNG `PngBuffer`.
///
/// PNG inputs pass through byte-for-byte (dimensions are extracted from the
/// IHDR chunk). Non-PNG inputs (JPEG/BMP/TIFF) require the Phase 2 decoders;
/// in the PNG-only baseline build they throw `ReportError::NotImplemented`.
///
/// Throws `ReportException`:
///   * CantOpenTemplate  — file missing / unreadable
///   * CantCopyDocxTemplate — corrupt PNG (bad signature, truncated IHDR)
///   * NotImplemented    — format detected but baseline build cannot decode
[[nodiscard]] PngBuffer load_as_png(const std::filesystem::path& path);

} // namespace docweft::docx
