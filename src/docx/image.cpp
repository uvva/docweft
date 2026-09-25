#include "docx/image.hpp"
#include "docweft/error.hpp"

#include <fmt/format.h>

#include <array>
#include <cstring>
#include <fstream>
#include <memory>
#include <vector>

// ── Optional decoders (wired up by Dependencies.cmake + CMakeLists.txt) ─────
#if defined(DOCWEFT_HAVE_STB)
#  define STB_IMAGE_IMPLEMENTATION
#  define STB_IMAGE_WRITE_IMPLEMENTATION
#  define STBI_NO_HDR
#  define STBI_NO_PIC
#  define STBI_NO_PNM
#  define STBI_NO_GIF
#  define STBI_NO_PSD
#  define STBI_NO_TGA
#  include <stb_image.h>
#  include <stb_image_write.h>
#endif

#if defined(DOCWEFT_HAVE_TIFF)
#  include <tiffio.h>
#endif

// ── Native-crash containment (Windows only) ──────────────────────────────
// stb_image and libtiff are third-party C decoders; malformed/corrupt input
// has been observed (via user crash dumps, not local repro — see
// DOCWEFT_UPSTREAM_FEEDBACK.md item 10) to trigger a genuine hardware
// exception (EXCEPTION_ACCESS_VIOLATION deep in stb_image's Huffman
// decoding) rather than a clean decode failure. That bypasses every
// try/catch a caller might have — it's a process-killing SEH, not a C++
// exception. _set_se_translator lets us convert it into a normal,
// catchable ReportException instead. This requires this translation unit
// to be compiled with /EHa (see CMakeLists.txt) — without it the CRT never
// invokes the translator and the hardware exception still kills the
// process.
#if defined(_MSC_VER)
#  include <eh.h>
#endif

namespace docweft::docx {

namespace {

constexpr std::array<unsigned char, 8> kPngSignature = {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};

constexpr std::array<unsigned char, 2> kJpegSignature  = {0xFF, 0xD8};
constexpr std::array<unsigned char, 2> kBmpSignature   = {0x42, 0x4D};  // "BM"
constexpr std::array<unsigned char, 4> kTiffLeSig      = {0x49, 0x49, 0x2A, 0x00};  // "II*\0"
constexpr std::array<unsigned char, 4> kTiffBeSig      = {0x4D, 0x4D, 0x00, 0x2A};  // "MM\0*"

std::string slurp(const std::filesystem::path& path) {
    if (!std::filesystem::exists(path)) {
        throw ReportException(
            ReportError::CantOpenTemplate,
            fmt::format("image file not found: {}", path.string()));
    }
    std::ifstream f(path, std::ios::binary);
    if (!f) {
        throw ReportException(
            ReportError::CantOpenTemplate,
            fmt::format("cannot open image file: {}", path.string()));
    }
    std::string bytes((std::istreambuf_iterator<char>(f)),
                       std::istreambuf_iterator<char>());
    if (bytes.empty()) {
        throw ReportException(
            ReportError::CantOpenTemplate,
            fmt::format("image file is empty: {}", path.string()));
    }
    return bytes;
}

template <std::size_t N>
bool starts_with(const std::string& bytes,
                 const std::array<unsigned char, N>& sig) {
    if (bytes.size() < N) return false;
    return std::memcmp(bytes.data(), sig.data(), N) == 0;
}

// Read a PNG's IHDR chunk to extract width/height. The PNG spec says the
// IHDR is always the first chunk and width/height are at byte offsets
// 16..19 and 20..23 respectively, big-endian uint32.
void parse_png_dimensions(const std::string& bytes,
                          std::uint32_t& width, std::uint32_t& height) {
    // Sig(8) + Len(4) + "IHDR"(4) + W(4) + H(4) = 24 bytes minimum to read H.
    if (bytes.size() < 24) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "PNG too small to contain IHDR");
    }
    if (std::memcmp(bytes.data() + 12, "IHDR", 4) != 0) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "PNG: first chunk is not IHDR");
    }
    auto read_be32 = [&bytes](std::size_t off) {
        return  (static_cast<std::uint32_t>(static_cast<unsigned char>(bytes[off    ])) << 24)
              | (static_cast<std::uint32_t>(static_cast<unsigned char>(bytes[off + 1])) << 16)
              | (static_cast<std::uint32_t>(static_cast<unsigned char>(bytes[off + 2])) <<  8)
              |  static_cast<std::uint32_t>(static_cast<unsigned char>(bytes[off + 3]));
    };
    width  = read_be32(16);
    height = read_be32(20);
    if (width == 0 || height == 0) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "PNG: zero width or height");
    }
}

} // namespace

ImageFormat detect_format(const std::filesystem::path& path) {
    if (!std::filesystem::exists(path)) return ImageFormat::Unknown;
    std::ifstream f(path, std::ios::binary);
    if (!f) return ImageFormat::Unknown;
    char head[8] = {};
    f.read(head, 8);
    const std::string bytes(head, static_cast<std::size_t>(f.gcount()));
    if (starts_with(bytes, kPngSignature)) return ImageFormat::Png;
    if (starts_with(bytes, kJpegSignature)) return ImageFormat::Jpeg;
    if (starts_with(bytes, kBmpSignature))  return ImageFormat::Bmp;
    if (starts_with(bytes, kTiffLeSig) ||
        starts_with(bytes, kTiffBeSig))     return ImageFormat::Tiff;
    return ImageFormat::Unknown;
}

// ── Decode → RGBA8 → re-encode as PNG (Phase 2 helpers) ────────────────────

namespace {

#if defined(DOCWEFT_HAVE_STB)

// Accumulator that stb_image_write can pump PNG bytes into via a callback,
// without dragging in FILE* or a temp file on disk.
struct StbSink {
    std::string bytes;
};
void stb_sink_write(void* ctx, void* data, int len) {
    auto* sink = static_cast<StbSink*>(ctx);
    sink->bytes.append(static_cast<const char*>(data),
                       static_cast<std::size_t>(len));
}

// Encode a raw RGBA8 pixel buffer (w × h) as PNG and return the bytes.
std::string encode_png_from_rgba(const unsigned char* rgba,
                                 int width, int height)
{
    StbSink sink;
    const int stride = width * 4;
    const int rc = stbi_write_png_to_func(
        &stb_sink_write, &sink, width, height, 4, rgba, stride);
    if (rc == 0 || sink.bytes.empty()) {
        throw ReportException(
            ReportError::SaveFailed,
            "stb_image_write failed to encode PNG");
    }
    return std::move(sink.bytes);
}

// Shared stb_image decode path for in-memory JPEG/BMP → RGBA8 → PNG.
PngBuffer decode_via_stb(const std::string& bytes, const char* format_label) {
    int w = 0, h = 0, ch = 0;
    unsigned char* raw = stbi_load_from_memory(
        reinterpret_cast<const unsigned char*>(bytes.data()),
        static_cast<int>(bytes.size()),
        &w, &h, &ch, 4 /* force RGBA */);
    if (!raw) {
        const char* why = stbi_failure_reason();
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("stb_image failed to decode {}: {}",
                        format_label, why ? why : "unknown"));
    }
    std::unique_ptr<unsigned char, decltype(&stbi_image_free)>
        raw_owner(raw, &stbi_image_free);

    PngBuffer out;
    out.width  = static_cast<std::uint32_t>(w);
    out.height = static_cast<std::uint32_t>(h);
    out.bytes  = encode_png_from_rgba(raw_owner.get(), w, h);
    return out;
}
#endif  // DOCWEFT_HAVE_STB

#if defined(DOCWEFT_HAVE_TIFF) && defined(DOCWEFT_HAVE_STB)

// In-memory TIFF source so libtiff can read from our std::string buffer
// without touching the filesystem twice (we already slurped it).
struct TiffMem {
    const unsigned char* data;
    toff_t               size;
    toff_t               pos;
};
tsize_t tiff_mem_read(thandle_t h, tdata_t buf, tsize_t len) {
    auto* m = static_cast<TiffMem*>(h);
    const toff_t avail = (m->pos < m->size) ? (m->size - m->pos) : 0;
    const toff_t take  = (static_cast<toff_t>(len) < avail) ? len : avail;
    std::memcpy(buf, m->data + m->pos, take);
    m->pos += take;
    return static_cast<tsize_t>(take);
}
tsize_t tiff_mem_write(thandle_t, tdata_t, tsize_t) { return 0; }  // read-only
toff_t  tiff_mem_seek(thandle_t h, toff_t off, int whence) {
    auto* m = static_cast<TiffMem*>(h);
    switch (whence) {
        case SEEK_SET: m->pos = off; break;
        case SEEK_CUR: m->pos += off; break;
        case SEEK_END: m->pos = m->size + off; break;
    }
    return m->pos;
}
int     tiff_mem_close(thandle_t) { return 0; }
toff_t  tiff_mem_filesize(thandle_t h) {
    return static_cast<TiffMem*>(h)->size;
}
int     tiff_mem_map(thandle_t, tdata_t*, toff_t*) { return 0; }
void    tiff_mem_unmap(thandle_t, tdata_t, toff_t) {}

PngBuffer decode_via_libtiff(const std::string& bytes) {
    TiffMem mem{reinterpret_cast<const unsigned char*>(bytes.data()),
                 static_cast<toff_t>(bytes.size()), 0};
    TIFF* tif = TIFFClientOpen(
        "docweft-memory", "r", &mem,
        tiff_mem_read, tiff_mem_write, tiff_mem_seek, tiff_mem_close,
        tiff_mem_filesize, tiff_mem_map, tiff_mem_unmap);
    if (!tif) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "libtiff failed to open TIFF stream");
    }
    std::unique_ptr<TIFF, decltype(&TIFFClose)> tif_owner(tif, &TIFFClose);

    std::uint32_t w = 0, h = 0;
    TIFFGetField(tif, TIFFTAG_IMAGEWIDTH,  &w);
    TIFFGetField(tif, TIFFTAG_IMAGELENGTH, &h);
    if (w == 0 || h == 0) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "TIFF: zero width or height");
    }
    // Cap to something Word can actually render.
    if (w > 32767u || h > 32767u) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "TIFF image too large (> 32767 px per side)");
    }

    std::vector<std::uint32_t> argb(static_cast<std::size_t>(w) * h);
    if (!TIFFReadRGBAImageOriented(tif, w, h, argb.data(),
                                    ORIENTATION_TOPLEFT, /*stopOnError=*/0)) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "libtiff TIFFReadRGBAImage failed");
    }
    // libtiff packs pixels as ABGR in a uint32 on little-endian; repack to RGBA
    // byte order for stb_image_write.
    std::vector<unsigned char> rgba(argb.size() * 4);
    for (std::size_t i = 0; i < argb.size(); ++i) {
        const std::uint32_t p = argb[i];
        rgba[i * 4 + 0] = static_cast<unsigned char>( p        & 0xFF); // R
        rgba[i * 4 + 1] = static_cast<unsigned char>((p >>  8) & 0xFF); // G
        rgba[i * 4 + 2] = static_cast<unsigned char>((p >> 16) & 0xFF); // B
        rgba[i * 4 + 3] = static_cast<unsigned char>((p >> 24) & 0xFF); // A
    }

    PngBuffer out;
    out.width  = w;
    out.height = h;
    out.bytes  = encode_png_from_rgba(rgba.data(),
                                      static_cast<int>(w),
                                      static_cast<int>(h));
    return out;
}

#endif  // DOCWEFT_HAVE_TIFF && DOCWEFT_HAVE_STB

[[noreturn]] void throw_missing_decoder(const std::filesystem::path& path,
                                        const char* fmt,
                                        const char* needed) {
    throw ReportException(
        ReportError::NotImplemented,
        fmt::format("cannot decode {} ({}): this build was compiled without "
                    "{}. Install the dependency or convert the file to PNG.",
                    path.string(), fmt, needed));
}

#if defined(_MSC_VER)
// Signature fixed by _se_translator_function; must not be [[noreturn]] or
// otherwise altered, or _set_se_translator won't accept it.
void seh_to_report_exception(unsigned int code, struct _EXCEPTION_POINTERS*) {
    throw ReportException(
        ReportError::CantCopyDocxTemplate,
        fmt::format("native exception 0x{:08X} while decoding image data — "
                    "the file is likely corrupt or malformed", code));
}

// RAII-scoped so only the decode call below runs under the translator, not
// the rest of the process — _set_se_translator is a per-thread global.
class ScopedSehTranslator {
public:
    ScopedSehTranslator()
        : prev_(_set_se_translator(&seh_to_report_exception)) {}
    ~ScopedSehTranslator() { _set_se_translator(prev_); }
    ScopedSehTranslator(const ScopedSehTranslator&) = delete;
    ScopedSehTranslator& operator=(const ScopedSehTranslator&) = delete;
private:
    _se_translator_function prev_;
};
#endif // _MSC_VER

} // namespace

PngBuffer load_as_png(const std::filesystem::path& path) {
    std::string bytes = slurp(path);

#if defined(_MSC_VER)
    // Guards the stb_image / libtiff decode calls below — see the
    // "Native-crash containment" comment near the top of this file.
    ScopedSehTranslator seh_guard;
#endif

    if (starts_with(bytes, kPngSignature)) {
        PngBuffer buf;
        parse_png_dimensions(bytes, buf.width, buf.height);
        buf.bytes = std::move(bytes);
        return buf;
    }

    if (starts_with(bytes, kJpegSignature)) {
#if defined(DOCWEFT_HAVE_STB)
        return decode_via_stb(bytes, "JPEG");
#else
        throw_missing_decoder(path, "JPEG", "stb_image");
#endif
    }

    if (starts_with(bytes, kBmpSignature)) {
#if defined(DOCWEFT_HAVE_STB)
        return decode_via_stb(bytes, "BMP");
#else
        throw_missing_decoder(path, "BMP", "stb_image");
#endif
    }

    if (starts_with(bytes, kTiffLeSig) || starts_with(bytes, kTiffBeSig)) {
#if defined(DOCWEFT_HAVE_TIFF) && defined(DOCWEFT_HAVE_STB)
        return decode_via_libtiff(bytes);
#else
        throw_missing_decoder(path, "TIFF", "libtiff + stb_image_write");
#endif
    }

    throw ReportException(
        ReportError::CantCopyDocxTemplate,
        fmt::format("unrecognised image format: {}", path.string()));
}

} // namespace docweft::docx
