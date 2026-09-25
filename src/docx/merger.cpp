#include "docx/merger.hpp"
#include "docx/image.hpp"
#include "docx/pdf_native.hpp"
#include "docweft/error.hpp"

#include <zip.h>

#include <fmt/format.h>

#include <algorithm>
#include <cctype>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <limits>
#include <regex>
#include <sstream>
#include <thread>
#include <vector>

#ifdef _WIN32
#  ifndef NOMINMAX
#    define NOMINMAX
#  endif
#  ifndef WIN32_LEAN_AND_MEAN
#    define WIN32_LEAN_AND_MEAN
#  endif
#  include <windows.h>
#else
#  include <csignal>
#  include <fcntl.h>
#  include <sys/wait.h>
#  include <unistd.h>
#endif

namespace docweft::docx {

namespace {

/// RAII helper around libzip's zip_t*.
class ZipFile {
public:
    explicit ZipFile(zip_t* z) noexcept : z_(z) {}
    ~ZipFile() { if (z_) zip_discard(z_); }
    ZipFile(const ZipFile&) = delete;
    ZipFile& operator=(const ZipFile&) = delete;

    zip_t*  get()   const noexcept { return z_; }
    zip_t*  release() noexcept { auto* t = z_; z_ = nullptr; return t; }

private:
    zip_t* z_;
};

/// Read every file from an open zip archive into a name→bytes map.
std::unordered_map<std::string, std::string>
slurp_archive(zip_t* z) {
    std::unordered_map<std::string, std::string> out;
    const zip_int64_t n = zip_get_num_entries(z, 0);
    if (n < 0) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "zip_get_num_entries failed");
    }

    for (zip_int64_t i = 0; i < n; ++i) {
        zip_stat_t st{};
        zip_stat_init(&st);
        if (zip_stat_index(z, static_cast<zip_uint64_t>(i), 0, &st) != 0) {
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  fmt::format("zip_stat_index({}) failed", i));
        }
        const std::string name = st.name ? st.name : "";
        // Skip directory entries
        if (!name.empty() && name.back() == '/') continue;

        zip_file_t* f = zip_fopen_index(z, static_cast<zip_uint64_t>(i), 0);
        if (!f) {
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  fmt::format("zip_fopen_index({}) failed", i));
        }

        std::string buf;
        buf.resize(static_cast<std::size_t>(st.size));
        zip_int64_t read = zip_fread(f, buf.data(), st.size);
        zip_fclose(f);

        if (read < 0 || static_cast<zip_uint64_t>(read) != st.size) {
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  fmt::format("zip_fread({}) short read", name));
        }

        out.emplace(std::move(const_cast<std::string&>(name)), std::move(buf));
    }
    return out;
}

} // namespace

// ────────────────────────────────────────────────────────────────────────────

DocxMerger::DocxMerger()  = default;
DocxMerger::~DocxMerger() = default;

void DocxMerger::setCodePage(const std::string& code_page) {
    // v1 supports UTF-8 only; this is effectively a no-op guard.
    if (code_page != "UTF-8") {
        throw ReportException(
            ReportError::NotImplemented,
            fmt::format("only UTF-8 is supported in v1 (got {:?})", code_page));
    }
    code_page_ = code_page;
}

// ── load / save ─────────────────────────────────────────────────────────────

void DocxMerger::read_archive(const std::filesystem::path& path) {
    if (!std::filesystem::exists(path)) {
        throw ReportException(
            ReportError::CantOpenTemplate,
            fmt::format("template not found: {}", path.string()));
    }

    int err = 0;
    zip_t* raw = zip_open(path.string().c_str(), ZIP_RDONLY, &err);
    if (!raw) {
        zip_error_t zerr;
        zip_error_init_with_code(&zerr, err);
        const std::string msg = fmt::format("zip_open failed: {}",
                                             zip_error_strerror(&zerr));
        zip_error_fini(&zerr);
        throw ReportException(ReportError::CantOpenTemplate, msg);
    }
    ZipFile zf{raw};

    parts_ = slurp_archive(zf.get());

    if (parts_.find(kDocumentPart) == parts_.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("missing required part: {}", kDocumentPart));
    }
}

void DocxMerger::load(const std::string& path) {
    parts_.clear();
    document_.reset();
    pasted_.clear();
    loaded_ = false;

    read_archive(std::filesystem::path(path));
    reparse_document();
    loaded_ = true;
}

void DocxMerger::reparse_document() {
    const std::string& xml = parts_.at(kDocumentPart);
    const auto result = document_.load_buffer(xml.data(), xml.size());
    if (!result) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("pugixml parse error: {}", result.description()));
    }
}

void DocxMerger::reserialize_document() {
    std::ostringstream oss;
    document_.save(oss, "", pugi::format_raw | pugi::format_no_declaration);
    // DOCX requires an XML declaration — pugixml emits one unless disabled.
    // We emit our own to preserve exactly the standalone attribute MS Word uses.
    std::string xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)";
    xml += oss.str();
    parts_[kDocumentPart] = std::move(xml);
}

void DocxMerger::write_archive(const std::filesystem::path& path) const {
    write_archive(parts_, path);
}

void DocxMerger::write_archive(
    const std::unordered_map<std::string, std::string>& parts,
    const std::filesystem::path&                         path)
{
    // Overwrite target (libzip opens append-mode otherwise).
    std::error_code ec;
    std::filesystem::remove(path, ec);

    int err = 0;
    zip_t* raw = zip_open(path.string().c_str(), ZIP_CREATE | ZIP_EXCL, &err);
    if (!raw) {
        zip_error_t zerr;
        zip_error_init_with_code(&zerr, err);
        const std::string msg = fmt::format("zip_open(create) failed: {}",
                                             zip_error_strerror(&zerr));
        zip_error_fini(&zerr);
        throw ReportException(ReportError::SaveFailed, msg);
    }

    // libzip's zip_source_buffer keeps a pointer into `data` until
    // zip_close() — we must keep all buffers alive until then.
    // Since `parts` is stable (std::unordered_map, not mutated during save),
    // we pass `freep=0` and rely on its lifetime.
    for (const auto& [name, bytes] : parts) {
        zip_source_t* src = zip_source_buffer(raw, bytes.data(), bytes.size(), 0);
        if (!src) {
            zip_discard(raw);
            throw ReportException(
                ReportError::SaveFailed,
                fmt::format("zip_source_buffer failed for {}", name));
        }
        const zip_int64_t idx = zip_file_add(raw, name.c_str(), src,
                                             ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8);
        if (idx < 0) {
            zip_source_free(src);
            const std::string msg = fmt::format("zip_file_add({}) failed: {}",
                                                 name, zip_strerror(raw));
            zip_discard(raw);
            throw ReportException(ReportError::SaveFailed, msg);
        }
    }

    if (zip_close(raw) != 0) {
        const std::string msg = fmt::format("zip_close failed: {}",
                                             zip_strerror(raw));
        zip_discard(raw);
        throw ReportException(ReportError::SaveFailed, msg);
    }
}

// ── Converter dispatch (Stage 6) ────────────────────────────────────────────
//
// `.docx → .pdf/.html` is tried through a chain of converters, in this
// order. Every converter that is present is tried; if one fails, the next
// one gets the document.
//
//   1. Microsoft Word (Windows and macOS) — the result matches what the user
//      sees in Word interactively.
//        Windows: detected via HKCR\Word.Application, driven by a generated
//                 VBScript through `cscript //B`.
//        macOS:   detected via Microsoft Word.app in /Applications or
//                 ~/Applications, driven by AppleScript through `osascript`.
//                 The first run shows the system "allow to control Microsoft
//                 Word" (Automation) prompt; the host app bundle needs
//                 NSAppleEventsUsageDescription (and, under hardened runtime,
//                 com.apple.security.automation.apple-events).
//   2. LibreOffice `soffice --headless --convert-to` (all platforms).
//      Detected via `DOCWEFT_SOFFICE` override, Program Files lookups
//      (Windows), LibreOffice.app bundle lookups (macOS), and finally a
//      `soffice --version` PATH probe.
//   3. Native PDF renderer (PoDoFo) — .pdf only; compiled in with
//      DOCWEFT_ENABLE_NATIVE_PDF=ON.
//   4. Remote converter — .pdf only; compiled in with
//      DOCWEFT_ENABLE_REMOTE_CONVERTER=ON and active only when
//      DOCWEFT_CONVERTER_URL is set (no default endpoint). Contract is
//      Gotenberg's LibreOffice route: multipart POST with the .docx in the
//      `files` field, PDF bytes in the response body. Sent with the `curl`
//      CLI.
//
// Env overrides:
//   DOCWEFT_DISABLE_CONVERTERS — any non-empty value disables the whole
//                                chain. Used by tests to assert the
//                                NoConverter path.
//   DOCWEFT_NO_MSWORD          — any non-empty value skips Word.
//   DOCWEFT_SOFFICE            — absolute path to soffice. If set non-empty
//                                it is used as-is (no other LibreOffice
//                                probe). Empty value disables the
//                                LibreOffice branch entirely.
//   DOCWEFT_SOFFICE_TIMEOUT    — LibreOffice time limit, seconds (default
//                                120). On expiry soffice and its child
//                                processes are killed and the chain moves on.
//   DOCWEFT_CONVERTER_URL      — remote converter endpoint, e.g.
//                                http://localhost:3000/forms/libreoffice/convert
//   DOCWEFT_CONVERTER_TOKEN    — optional; sent as `Authorization: Bearer`.
//                                Never included in error messages.
//   DOCWEFT_CONVERTER_TIMEOUT  — remote request timeout, seconds (default 120).
//
// If nothing in the chain is available, save() throws `NoConverter` listing
// why each converter was skipped; if everything available failed, it throws
// `SaveFailed` listing each failure.

namespace {

enum class ConverterKind { MSWord, LibreOffice, NativePdf, Remote };

struct Converter {
    ConverterKind kind;
    std::string   location;  // soffice path, Word.app path or endpoint URL
};

const char* converter_label(ConverterKind kind) {
    switch (kind) {
        case ConverterKind::MSWord:      return "Microsoft Word";
        case ConverterKind::LibreOffice: return "LibreOffice";
        case ConverterKind::NativePdf:   return "native PDF renderer";
        case ConverterKind::Remote:      return "remote converter";
    }
    return "unknown";
}

[[nodiscard]] bool env_nonempty(const char* name) {
    const char* v = std::getenv(name);
    return v != nullptr && *v != '\0';
}

// Positive number of seconds from env var `name`, or `fallback` when unset
// or not a number.
[[nodiscard]] int env_seconds(const char* name, int fallback) {
    const char* v = std::getenv(name);
    if (v == nullptr) return fallback;
    try { return std::max(1, std::stoi(v)); } catch (...) { return fallback; }
}

#ifdef __APPLE__
[[nodiscard]] std::string home_dir() {
    const char* home = std::getenv("HOME");
    return (home != nullptr) ? home : "";
}
#endif

// Word install location, or empty when Word is not installed.
[[nodiscard]] std::string find_msword() {
#if defined(_WIN32)
    // HKCR\Word.Application exists iff some Office version is installed and
    // has registered its COM class. `reg query` is present in every Windows
    // install and exits 0 on hit, 1 on miss.
    if (std::system("reg query HKCR\\Word.Application >nul 2>nul") == 0) {
        return "Word.Application";
    }
#elif defined(__APPLE__)
    std::vector<std::string> candidates = {"/Applications/Microsoft Word.app"};
    if (const std::string home = home_dir(); !home.empty()) {
        candidates.push_back(home + "/Applications/Microsoft Word.app");
    }
    for (const auto& c : candidates) {
        if (std::filesystem::exists(c)) return c;
    }
#endif
    return {};
}

// soffice to run, or empty when LibreOffice is not found / disabled.
[[nodiscard]] std::string find_soffice() {
    const char* const override_path = std::getenv("DOCWEFT_SOFFICE");
    if (override_path != nullptr) {
        if (*override_path == '\0') return {};  // explicit opt-out
        if (std::filesystem::exists(override_path)) return override_path;
        return {};
    }

    std::vector<std::string> candidates;
#if defined(_WIN32)
    candidates = {
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    };
#elif defined(__APPLE__)
    // The official .dmg (and the Homebrew cask) install an app bundle and
    // never touch PATH; apps started from Finder/Dock also get launchd's
    // minimal PATH, so the probe below alone would miss a normal install.
    candidates.emplace_back("/Applications/LibreOffice.app/Contents/MacOS/soffice");
    if (const std::string home = home_dir(); !home.empty()) {
        candidates.push_back(home + "/Applications/LibreOffice.app/Contents/MacOS/soffice");
    }
#endif
    for (const auto& c : candidates) {
        if (std::filesystem::exists(c)) return c;
    }

#ifdef _WIN32
    const char* const probe = "soffice --version >nul 2>nul";
#else
    const char* const probe = "soffice --version >/dev/null 2>&1";
#endif
    if (std::system(probe) == 0) return "soffice";
    return {};
}

struct ConverterChain {
    std::vector<Converter>   available;  // in the order they will be tried
    std::vector<std::string> skipped;    // "<label>: <why>" for the rest
};

// Build the chain for `format` ("pdf" or "html").
ConverterChain find_converters(const std::string& format) {
    (void)format;  // only consulted by the .pdf-only backends, if compiled in
    ConverterChain chain;
    if (env_nonempty("DOCWEFT_DISABLE_CONVERTERS")) {
        chain.skipped.emplace_back(
            "all converters: disabled by DOCWEFT_DISABLE_CONVERTERS");
        return chain;
    }
    auto skip = [&](ConverterKind kind, const std::string& why) {
        chain.skipped.push_back(fmt::format("{}: {}", converter_label(kind), why));
    };

    if (env_nonempty("DOCWEFT_NO_MSWORD")) {
        skip(ConverterKind::MSWord, "disabled by DOCWEFT_NO_MSWORD");
    } else if (std::string word = find_msword(); !word.empty()) {
        chain.available.push_back({ConverterKind::MSWord, std::move(word)});
    } else {
#if defined(_WIN32) || defined(__APPLE__)
        skip(ConverterKind::MSWord, "not installed");
#else
        skip(ConverterKind::MSWord, "not available on this platform");
#endif
    }

    if (std::string soffice = find_soffice(); !soffice.empty()) {
        chain.available.push_back({ConverterKind::LibreOffice, std::move(soffice)});
    } else if (const char* o = std::getenv("DOCWEFT_SOFFICE"); o != nullptr) {
        skip(ConverterKind::LibreOffice,
             *o == '\0' ? "disabled by empty DOCWEFT_SOFFICE"
                        : "DOCWEFT_SOFFICE points to a missing file");
    } else {
        skip(ConverterKind::LibreOffice, "not installed");
    }

#if defined(DOCWEFT_HAVE_PODOFO)
    if (format == "pdf") {
        chain.available.push_back({ConverterKind::NativePdf, {}});
    } else {
        skip(ConverterKind::NativePdf, "produces .pdf only");
    }
#else
    skip(ConverterKind::NativePdf, "not built (DOCWEFT_ENABLE_NATIVE_PDF=OFF)");
#endif

#if defined(DOCWEFT_HAVE_REMOTE_CONVERTER)
    if (format != "pdf") {
        skip(ConverterKind::Remote, "produces .pdf only");
    } else if (!env_nonempty("DOCWEFT_CONVERTER_URL")) {
        skip(ConverterKind::Remote, "DOCWEFT_CONVERTER_URL is not set");
    } else {
        chain.available.push_back(
            {ConverterKind::Remote, std::getenv("DOCWEFT_CONVERTER_URL")});
    }
#else
    skip(ConverterKind::Remote, "not built (DOCWEFT_ENABLE_REMOTE_CONVERTER=OFF)");
#endif

    return chain;
}

// ── Default theme for converter input ───────────────────────────────────────
//
// Chart series without an explicit <c:spPr> take their "automatic" colors
// from the document theme (accent1..6). Word falls back to its built-in
// Office theme when the package has none; LibreOffice does not and draws
// such series with no fill at all — the PDF/HTML gets axes and labels but
// an empty plot area. Word-authored templates always ship a theme, so this
// only bites hand-generated ones (e.g. examples/generate_template.cpp).
//
// The fix is applied to the scratch .docx handed to the converter only;
// parts_ — and therefore a later save(".docx") — is left untouched.

constexpr const char* kThemeRelType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
constexpr const char* kThemeContentType =
    "application/vnd.openxmlformats-officedocument.theme+xml";

// Standard Office 2013+ color scheme; fonts/format scheme kept minimal.
constexpr const char* kDefaultThemeXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
    R"(<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">)"
    R"(<a:themeElements>)"
    R"(<a:clrScheme name="Office">)"
    R"(<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>)"
    R"(<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>)"
    R"(<a:dk2><a:srgbClr val="44546A"/></a:dk2>)"
    R"(<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>)"
    R"(<a:accent1><a:srgbClr val="4472C4"/></a:accent1>)"
    R"(<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>)"
    R"(<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>)"
    R"(<a:accent4><a:srgbClr val="FFC000"/></a:accent4>)"
    R"(<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>)"
    R"(<a:accent6><a:srgbClr val="70AD47"/></a:accent6>)"
    R"(<a:hlink><a:srgbClr val="0563C1"/></a:hlink>)"
    R"(<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>)"
    R"(</a:clrScheme>)"
    R"(<a:fontScheme name="Office">)"
    R"(<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>)"
    R"(<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>)"
    R"(</a:fontScheme>)"
    R"(<a:fmtScheme name="Office">)"
    R"(<a:fillStyleLst>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(</a:fillStyleLst>)"
    R"(<a:lnStyleLst>)"
    R"(<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
    R"(<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
    R"(<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
    R"(</a:lnStyleLst>)"
    R"(<a:effectStyleLst>)"
    R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
    R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
    R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
    R"(</a:effectStyleLst>)"
    R"(<a:bgFillStyleLst>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
    R"(</a:bgFillStyleLst>)"
    R"(</a:fmtScheme>)"
    R"(</a:themeElements>)"
    R"(</a:theme>)";

// True when the package has at least one chart part but its main document
// declares no theme relationship.
[[nodiscard]] bool needs_default_theme(
    const std::unordered_map<std::string, std::string>& parts)
{
    const bool has_chart = std::any_of(parts.begin(), parts.end(),
        [](const auto& kv) { return kv.first.rfind("word/charts/chart", 0) == 0; });
    if (!has_chart) return false;

    const auto rels_it = parts.find("word/_rels/document.xml.rels");
    if (rels_it == parts.end()) return false;  // no rels → charts can't be linked anyway
    pugi::xml_document rels;
    if (!rels.load_buffer(rels_it->second.data(), rels_it->second.size())) return false;
    for (auto rel : rels.child("Relationships").children("Relationship")) {
        if (std::strcmp(rel.attribute("Type").value(), kThemeRelType) == 0) return false;
    }
    return parts.find("[Content_Types].xml") != parts.end();
}

// Add word/theme/themeN.xml with kDefaultThemeXml, its relationship and its
// content-type override. Call only when needs_default_theme(parts) is true.
void add_default_theme(std::unordered_map<std::string, std::string>& parts) {
    std::string part_name;
    for (int n = 1; part_name.empty() || parts.count(part_name) != 0; ++n) {
        part_name = fmt::format("word/theme/theme{}.xml", n);
    }
    parts[part_name] = kDefaultThemeXml;

    auto& rels_xml = parts["word/_rels/document.xml.rels"];
    pugi::xml_document rels;
    rels.load_buffer(rels_xml.data(), rels_xml.size());
    auto rel = rels.child("Relationships").append_child("Relationship");
    rel.append_attribute("Id").set_value("rIdDocWeftTheme");
    rel.append_attribute("Type").set_value(kThemeRelType);
    rel.append_attribute("Target").set_value(part_name.substr(std::strlen("word/")).c_str());
    std::ostringstream rels_out;
    rels.save(rels_out, "", pugi::format_raw);
    rels_xml = rels_out.str();

    auto& ct_xml = parts["[Content_Types].xml"];
    pugi::xml_document ct;
    ct.load_buffer(ct_xml.data(), ct_xml.size());
    auto ovr = ct.child("Types").append_child("Override");
    ovr.append_attribute("PartName").set_value(("/" + part_name).c_str());
    ovr.append_attribute("ContentType").set_value(kThemeContentType);
    std::ostringstream ct_out;
    ct.save(ct_out, "", pugi::format_raw);
    ct_xml = ct_out.str();
}

#if defined(_WIN32) || defined(__APPLE__) || defined(DOCWEFT_HAVE_REMOTE_CONVERTER)
// Minimal shell-safe quoting. Rejects embedded double quotes rather than
// trying to escape them portably (cmd.exe and POSIX sh have different rules).
std::string shell_quote(const std::string& s) {
    if (s.find('"') != std::string::npos) {
        throw ReportException(
            ReportError::SaveFailed,
            fmt::format("path contains a double-quote character, refusing: {}", s));
    }
    return "\"" + s + "\"";
}
#endif

// Read the entirety of a file into a string, in binary mode.
std::string read_file_bytes(const std::filesystem::path& path) {
    std::ifstream in(path, std::ios::binary);
    return std::string((std::istreambuf_iterator<char>(in)),
                        std::istreambuf_iterator<char>());
}

// Captured stderr of a converter, trimmed, for error messages.
std::string read_error_text(const std::filesystem::path& path) {
    std::string text = read_file_bytes(path);
    while (!text.empty() && std::isspace(static_cast<unsigned char>(text.back()))) {
        text.pop_back();
    }
    constexpr std::size_t kMaxLen = 500;
    if (text.size() > kMaxLen) text = text.substr(0, kMaxLen) + "...";
    return text;
}

// Outcome of run_with_timeout().
struct ProcessResult {
    bool started   = false;
    bool timed_out = false;
    int  exit_code = -1;
};

// Run `args` (args[0] is looked up in PATH) with stdin/stdout on the null
// device and stderr into `stderr_file`, killing it — together with every
// process it started — once `timeout` expires. Arguments are passed as
// paths so they reach the OS in its native encoding (UTF-16 on Windows).
ProcessResult run_with_timeout(const std::vector<std::filesystem::path>& args,
                               const std::filesystem::path&              stderr_file,
                               std::chrono::seconds                      timeout)
{
    ProcessResult result;
#ifdef _WIN32
    // CommandLineToArgvW quoting: wrap in quotes, double the backslashes
    // that precede a quote, escape the quote itself.
    std::wstring cmdline;
    for (const auto& arg : args) {
        if (!cmdline.empty()) cmdline += L' ';
        cmdline += L'"';
        std::size_t backslashes = 0;
        for (const wchar_t c : arg.wstring()) {
            if (c == L'\\') { ++backslashes; continue; }
            if (c == L'"') cmdline.append(backslashes * 2 + 1, L'\\');
            else           cmdline.append(backslashes, L'\\');
            backslashes = 0;
            cmdline += c;
        }
        cmdline.append(backslashes * 2, L'\\');
        cmdline += L'"';
    }

    SECURITY_ATTRIBUTES inherit{sizeof(SECURITY_ATTRIBUTES), nullptr, TRUE};
    HANDLE null_in  = CreateFileW(L"NUL", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE,
                                  &inherit, OPEN_EXISTING, 0, nullptr);
    HANDLE err_out  = CreateFileW(stderr_file.wstring().c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                  &inherit, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    // A job object lets a timeout kill soffice.exe and the soffice.bin it
    // spawns in one go.
    HANDLE job = CreateJobObjectW(nullptr, nullptr);
    if (job != nullptr) {
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION limits{};
        limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        SetInformationJobObject(job, JobObjectExtendedLimitInformation, &limits, sizeof(limits));
    }

    STARTUPINFOW si{};
    si.cb         = sizeof(si);
    si.dwFlags    = STARTF_USESTDHANDLES;
    si.hStdInput  = null_in;
    si.hStdOutput = null_in;
    si.hStdError  = (err_out != INVALID_HANDLE_VALUE) ? err_out : null_in;
    PROCESS_INFORMATION pi{};
    if (CreateProcessW(nullptr, cmdline.data(), nullptr, nullptr, TRUE,
                       CREATE_SUSPENDED | CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        result.started = true;
        if (job != nullptr) AssignProcessToJobObject(job, pi.hProcess);
        ResumeThread(pi.hThread);
        const auto ms = static_cast<DWORD>(
            std::chrono::duration_cast<std::chrono::milliseconds>(timeout).count());
        if (WaitForSingleObject(pi.hProcess, ms) == WAIT_TIMEOUT) {
            result.timed_out = true;
            if (job != nullptr) TerminateJobObject(job, 1);
            else                TerminateProcess(pi.hProcess, 1);
            WaitForSingleObject(pi.hProcess, INFINITE);
        }
        DWORD code = 0;
        GetExitCodeProcess(pi.hProcess, &code);
        result.exit_code = static_cast<int>(code);
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }
    if (job != nullptr) CloseHandle(job);
    if (err_out != INVALID_HANDLE_VALUE) CloseHandle(err_out);
    if (null_in != INVALID_HANDLE_VALUE) CloseHandle(null_in);
#else
    // Everything the child needs is prepared before fork(): only
    // async-signal-safe calls are allowed between fork() and exec().
    std::vector<std::string> storage;
    storage.reserve(args.size());
    for (const auto& a : args) storage.push_back(a.string());
    std::vector<char*> argv;
    for (auto& a : storage) argv.push_back(a.data());
    argv.push_back(nullptr);
    const std::string err_path = stderr_file.string();

    const pid_t pid = fork();
    if (pid < 0) return result;
    if (pid == 0) {
        // Own process group, so a timeout can kill soffice and the
        // soffice.bin it starts together.
        setpgid(0, 0);
        const int null_fd = open("/dev/null", O_RDWR);
        const int err_fd  = open(err_path.c_str(), O_WRONLY | O_CREAT | O_TRUNC, 0600);
        if (null_fd >= 0) { dup2(null_fd, 0); dup2(null_fd, 1); }
        if (err_fd >= 0) dup2(err_fd, 2);
        execvp(argv[0], argv.data());
        _exit(127);
    }
    setpgid(pid, pid);  // also from the parent: no race with kill() below
    result.started = true;

    const auto deadline = std::chrono::steady_clock::now() + timeout;
    int status = 0;
    for (;;) {
        const pid_t done = waitpid(pid, &status, WNOHANG);
        if (done == pid) break;
        if (done < 0) return result;
        if (std::chrono::steady_clock::now() >= deadline) {
            result.timed_out = true;
            kill(-pid, SIGKILL);
            waitpid(pid, &status, 0);
            break;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
    }
    if (!result.timed_out) {
        result.exit_code = WIFEXITED(status) ? WEXITSTATUS(status) : 128 + WTERMSIG(status);
    }
#endif
    return result;
}

// Run `soffice --headless --convert-to <format> --outdir <outdir> <input>`.
// Produces `<outdir>/<input-stem>.<format>`. Returns an empty string on
// success, otherwise the reason. Bounded by DOCWEFT_SOFFICE_TIMEOUT —
// soffice can hang on a dialog or a stuck profile lock, and save() must not
// hang with it.
std::string run_soffice_conversion(const std::string& soffice,
                                   const std::string& format,  // "pdf" or "html"
                                   const std::filesystem::path& input,
                                   const std::filesystem::path& outdir)
{
    const int timeout_s = env_seconds("DOCWEFT_SOFFICE_TIMEOUT", 120);
    const std::filesystem::path err = outdir / "tf_soffice.err";
    const ProcessResult r = run_with_timeout(
        {soffice, "--headless", "--convert-to", format, "--outdir", outdir, input},
        err, std::chrono::seconds(timeout_s));

    if (!r.started) return "could not start soffice";
    if (r.timed_out) {
        return fmt::format("soffice did not finish within {} s "
                           "(DOCWEFT_SOFFICE_TIMEOUT) and was stopped", timeout_s);
    }
    if (r.exit_code != 0) {
        const std::string text = read_error_text(err);
        return text.empty()
            ? fmt::format("soffice exited with code {}", r.exit_code)
            : fmt::format("soffice exited with code {}: {}", r.exit_code, text);
    }
    return {};
}

#ifdef _WIN32
// Run Word via a generated VBScript. Exit codes:
//   0 — ok
//   2 — CreateObject("Word.Application") failed (unlikely after reg probe)
//   3 — Documents.Open failed (corrupted .docx, locked, etc.)
//   4 — SaveAs2 failed (unknown format, read-only target)
// cscript is always present on Windows 10/11 (no PowerShell ExecutionPolicy
// headaches) and can run in //B (batch, suppress dialogs) mode.
bool run_msword_conversion(const std::string& format,  // "pdf" or "html"
                           const std::filesystem::path& input,
                           const std::filesystem::path& output)
{
    // wdFormatPDF = 17, wdFormatFilteredHTML = 10 (both stable since Word 2007)
    const int fmt_code = (format == "pdf") ? 17 : 10;
    const std::filesystem::path vbs = output.parent_path() / "tf_msword.vbs";
    {
        std::ofstream os(vbs, std::ios::binary);
        if (!os) return false;
        // CRLF is what cscript expects; using binary mode so ofstream doesn't
        // double-convert on Windows.
        os << "Option Explicit\r\n"
              "Dim word, doc\r\n"
              "On Error Resume Next\r\n"
              "Set word = CreateObject(\"Word.Application\")\r\n"
              "If Err.Number <> 0 Then WScript.Quit 2\r\n"
              "word.Visible = False\r\n"
              "word.DisplayAlerts = 0\r\n"
              "Set doc = word.Documents.Open(WScript.Arguments(0), False, True)\r\n"
              "If Err.Number <> 0 Then\r\n"
              "    word.Quit\r\n"
              "    WScript.Quit 3\r\n"
              "End If\r\n"
              "doc.SaveAs2 WScript.Arguments(1), " << fmt_code << "\r\n"
              "If Err.Number <> 0 Then\r\n"
              "    doc.Close False\r\n"
              "    word.Quit\r\n"
              "    WScript.Quit 4\r\n"
              "End If\r\n"
              "doc.Close False\r\n"
              "word.Quit\r\n"
              "WScript.Quit 0\r\n";
    }
    const std::string cmd = fmt::format(
        "cscript //B //Nologo {} {} {} >nul 2>nul",
        shell_quote(vbs.string()),
        shell_quote(std::filesystem::absolute(input).string()),
        shell_quote(std::filesystem::absolute(output).string()));
    const int rc = std::system(cmd.c_str());
    std::error_code ec;
    std::filesystem::remove(vbs, ec);
    return rc == 0;
}
#endif  // _WIN32

// Unique per-call temp dir under the OS temp location.
std::filesystem::path make_scratch_dir() {
    const auto ts = std::chrono::duration_cast<std::chrono::nanoseconds>(
        std::chrono::steady_clock::now().time_since_epoch()).count();
    auto dir = std::filesystem::temp_directory_path() /
               fmt::format("docweft-{}", ts);
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    return dir;
}

// Encode raw bytes as base64 (RFC 4648, standard alphabet, padded).
std::string base64_encode(const std::string& data) {
    static constexpr char table[] =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        "abcdefghijklmnopqrstuvwxyz"
        "0123456789+/";

    std::string out;
    out.reserve(((data.size() + 2) / 3) * 4);

    std::size_t i = 0;
    const std::size_t n = data.size();
    for (; i + 3 <= n; i += 3) {
        const std::uint32_t chunk =
            (static_cast<std::uint8_t>(data[i]) << 16) |
            (static_cast<std::uint8_t>(data[i + 1]) << 8) |
            static_cast<std::uint8_t>(data[i + 2]);
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.push_back(table[(chunk >> 6) & 0x3F]);
        out.push_back(table[chunk & 0x3F]);
    }

    const std::size_t remaining = n - i;
    if (remaining == 1) {
        const std::uint32_t chunk = static_cast<std::uint32_t>(
            static_cast<std::uint8_t>(data[i])) << 16;
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.push_back('=');
        out.push_back('=');
    } else if (remaining == 2) {
        const std::uint32_t chunk =
            (static_cast<std::uint32_t>(static_cast<std::uint8_t>(data[i])) << 16) |
            (static_cast<std::uint32_t>(static_cast<std::uint8_t>(data[i + 1])) << 8);
        out.push_back(table[(chunk >> 18) & 0x3F]);
        out.push_back(table[(chunk >> 12) & 0x3F]);
        out.push_back(table[(chunk >> 6) & 0x3F]);
        out.push_back('=');
    }
    return out;
}

// MIME type to use when embedding `bytes` (the contents of `path`) as a
// data: URI. `image.hpp`'s detect_format() only recognizes the formats
// DocWeft can re-encode for DOCX embedding (PNG/JPEG/BMP/TIFF); it
// doesn't know GIF, SVG or WebP, all of which LibreOffice's HTML/chart
// export can still produce as sibling files. Sniff those separately by
// magic bytes before falling back to the file extension.
std::string sniff_data_uri_mime(const std::filesystem::path& path,
                                 const std::string& bytes) {
    switch (detect_format(path)) {
        case ImageFormat::Png:  return "image/png";
        case ImageFormat::Jpeg: return "image/jpeg";
        case ImageFormat::Bmp:  return "image/bmp";
        case ImageFormat::Tiff: return "image/tiff";
        default: break;
    }

    if (bytes.size() >= 6 &&
        (bytes.compare(0, 6, "GIF87a") == 0 || bytes.compare(0, 6, "GIF89a") == 0)) {
        return "image/gif";
    }
    if (bytes.size() >= 12 && bytes.compare(0, 4, "RIFF") == 0 &&
        bytes.compare(8, 4, "WEBP") == 0) {
        return "image/webp";
    }

    std::string ext = path.extension().string();
    std::transform(ext.begin(), ext.end(), ext.begin(),
                    [](unsigned char c) { return std::tolower(c); });
    if (ext == ".svg") return "image/svg+xml";
    if (ext == ".png") return "image/png";
    if (ext == ".jpg" || ext == ".jpeg") return "image/jpeg";
    if (ext == ".gif") return "image/gif";
    if (ext == ".bmp") return "image/bmp";
    if (ext == ".webp") return "image/webp";

    return "application/octet-stream";
}

// LibreOffice's HTML export writes referenced images as sibling files next
// to `html_path` (e.g. "output_html_....png") instead of embedding them —
// <img src="..."> in the produced HTML points at those bare filenames.
// Rewrite `html_path` in place so every such reference becomes an inline
// base64 data: URI, making the HTML self-contained. `siblings` is consumed
// only for its bytes; callers are still responsible for deleting the files.
void inline_sibling_images(const std::filesystem::path& html_path,
                            const std::vector<std::filesystem::path>& siblings) {
    if (siblings.empty()) return;

    std::string html = read_file_bytes(html_path);

    for (const auto& sibling : siblings) {
        const std::string bytes = read_file_bytes(sibling);
        const std::string data_uri =
            fmt::format("data:{};base64,{}",
                        sniff_data_uri_mime(sibling, bytes),
                        base64_encode(bytes));

        const std::string needle = "\"" + sibling.filename().string() + "\"";
        const std::string replacement = "\"" + data_uri + "\"";
        std::size_t pos = 0;
        while ((pos = html.find(needle, pos)) != std::string::npos) {
            html.replace(pos, needle.size(), replacement);
            pos += replacement.size();
        }
    }

    std::ofstream out(html_path, std::ios::binary | std::ios::trunc);
    out.write(html.data(), static_cast<std::streamsize>(html.size()));
}


#ifdef __APPLE__
// Run Word for Mac via AppleScript. Returns an empty string on success,
// otherwise the reason. Word is quit afterwards only if this call started
// it. `with timeout` bounds each Apple Event, so a dialog nobody answers
// (sign-in, "Grant File Access" for the sandbox) fails the step instead of
// hanging save() forever.
std::string run_msword_mac_conversion(const std::string& format,  // "pdf" or "html"
                                      const std::filesystem::path& input,
                                      const std::filesystem::path& output)
{
    const std::filesystem::path dir    = output.parent_path();
    const std::filesystem::path script = dir / "tf_msword.applescript";
    const std::filesystem::path err    = dir / "tf_msword.err";
    {
        std::ofstream os(script, std::ios::binary);
        if (!os) return "cannot write the AppleScript file";
        os << "on run argv\n"
              "  set inFile to POSIX file (item 1 of argv)\n"
              "  set outPath to item 2 of argv\n"
              "  set wasRunning to application \"Microsoft Word\" is running\n"
              "  with timeout of 300 seconds\n"
              "    tell application \"Microsoft Word\"\n"
              "      open inFile\n"
              "      set doc to active document\n"
              "      try\n"
              "        save as doc file name outPath file format "
           << (format == "pdf" ? "format PDF" : "format filtered HTML") << "\n"
              "      on error errMsg number errNum\n"
              "        close doc saving no\n"
              "        if not wasRunning then quit saving no\n"
              "        error errMsg number errNum\n"
              "      end try\n"
              "      close doc saving no\n"
              "      if not wasRunning then quit saving no\n"
              "    end tell\n"
              "  end timeout\n"
              "end run\n";
    }
    const std::string cmd = fmt::format(
        "osascript {} {} {} >/dev/null 2>{}",
        shell_quote(script.string()),
        shell_quote(std::filesystem::absolute(input).string()),
        shell_quote(std::filesystem::absolute(output).string()),
        shell_quote(err.string()));
    const int rc = std::system(cmd.c_str());
    std::error_code ec;
    std::filesystem::remove(script, ec);
    if (rc == 0) return {};

    const std::string text = read_error_text(err);
    if (text.find("-1743") != std::string::npos) {
        return "macOS denied control of Microsoft Word — allow it in System "
               "Settings → Privacy & Security → Automation";
    }
    return text.empty() ? "osascript failed" : "osascript: " + text;
}
#endif  // __APPLE__

#if defined(DOCWEFT_HAVE_REMOTE_CONVERTER)
// Quote a value for a curl config file ("..." with backslash escapes).
std::string curl_config_quote(const std::string& s) {
    std::string out = "\"";
    for (const char c : s) {
        if (c == '"' || c == '\\') out.push_back('\\');
        out.push_back(c);
    }
    out.push_back('"');
    return out;
}

// POST `input` to `url` as multipart `files` (Gotenberg's LibreOffice route
// contract) and store the response body in `output`. Returns an empty string
// on success, otherwise the reason. All options go through a curl config
// file so the auth token never shows up on a command line (visible in the
// process list) — and it is never put into the returned text.
std::string run_remote_conversion(const std::string& url,
                                  const std::filesystem::path& input,
                                  const std::filesystem::path& output)
{
    const int timeout_s = env_seconds("DOCWEFT_CONVERTER_TIMEOUT", 120);

    const std::filesystem::path dir = output.parent_path();
    const std::filesystem::path cfg = dir / "tf_curl.cfg";
    const std::filesystem::path err = dir / "tf_curl.err";
    {
        std::ofstream os(cfg, std::ios::binary);
        if (!os) return "cannot write the curl config file";
        os << "silent\nshow-error\nfail\n"
           << "max-time = " << timeout_s << "\n"
           << "url = " << curl_config_quote(url) << "\n"
           << "form = " << curl_config_quote("files=@" + input.string()) << "\n"
           << "output = " << curl_config_quote(output.string()) << "\n";
        if (const char* token = std::getenv("DOCWEFT_CONVERTER_TOKEN");
            token != nullptr && *token != '\0') {
            os << "header = "
               << curl_config_quote(std::string("Authorization: Bearer ") + token) << "\n";
        }
    }
#ifdef _WIN32
    constexpr const char* kNull = "nul";
#else
    constexpr const char* kNull = "/dev/null";
#endif
    const std::string cmd = fmt::format("curl --config {} >{} 2>{}",
                                        shell_quote(cfg.string()), kNull,
                                        shell_quote(err.string()));
    const int rc = std::system(cmd.c_str());
    std::error_code ec;
    std::filesystem::remove(cfg, ec);

    if (rc != 0) {
        const std::string text = read_error_text(err);
        return text.empty() ? "curl failed" : text;
    }
    // A misconfigured endpoint (wrong route, HTML error page with 200)
    // must fail loudly rather than be saved as report.pdf.
    std::ifstream in(output, std::ios::binary);
    char magic[4] = {};
    in.read(magic, 4);
    if (in.gcount() != 4 || std::string(magic, 4) != "%PDF") {
        return "endpoint response is not a PDF";
    }
    return {};
}
#endif  // DOCWEFT_HAVE_REMOTE_CONVERTER

} // namespace

void DocxMerger::save(const std::string& path) {
    if (!loaded_) {
        throw ReportException(ReportError::SaveFailed,
                              "save() called before load()");
    }

    const std::filesystem::path p(path);
    const std::string ext = p.extension().string();

    if (ext == ".pdf" || ext == ".html" || ext == ".htm") {
        const std::string conv_format = (ext == ".pdf") ? "pdf" : "html";
        const ConverterChain chain = find_converters(conv_format);
        if (chain.available.empty()) {
            std::string why;
            for (const auto& s : chain.skipped) why += "\n  - " + s;
            throw ReportException(
                ReportError::NoConverter,
                fmt::format("No conversion tool available to produce {}:{}",
                            ext, why));
        }

        // Two-step: serialize to .docx in a scratch dir, then convert.
        const std::filesystem::path scratch = make_scratch_dir();
        const std::filesystem::path scratch_docx = scratch / "output.docx";

        std::error_code cleanup_ec;
        auto cleanup = [&]() {
            std::filesystem::remove_all(scratch, cleanup_ec);
        };

        reserialize_document();
        try {
            // Converter input only — see add_default_theme().
            std::unordered_map<std::string, std::string> themed;
            if (needs_default_theme(parts_)) {
                themed = parts_;
                add_default_theme(themed);
            }
            write_archive(themed.empty() ? parts_ : themed, scratch_docx);
        } catch (...) {
            cleanup();
            throw;
        }

        // Try each converter in turn; each gets its own output directory so
        // leftovers of a failed attempt can't be mistaken for output.
        struct Failure {
            ReportError code;
            std::string text;
        };
        std::vector<Failure> failures;
        std::filesystem::path produced;
        for (std::size_t i = 0; i < chain.available.size() && produced.empty(); ++i) {
            const Converter& conv = chain.available[i];
            const std::filesystem::path outdir = scratch / fmt::format("try{}", i);
            const std::filesystem::path target = outdir / ("output." + conv_format);
            std::error_code dir_ec;
            std::filesystem::create_directories(outdir, dir_ec);

            std::string error;
            ReportError code = ReportError::SaveFailed;
            try {
                switch (conv.kind) {
                    case ConverterKind::MSWord:
#if defined(_WIN32)
                        if (!run_msword_conversion(conv_format, scratch_docx, target)) {
                            error = "conversion failed";
                        }
#elif defined(__APPLE__)
                        error = run_msword_mac_conversion(conv_format, scratch_docx, target);
#endif
                        break;
                    case ConverterKind::LibreOffice:
                        error = run_soffice_conversion(conv.location, conv_format,
                                                       scratch_docx, outdir);
                        break;
                    case ConverterKind::NativePdf:
#if defined(DOCWEFT_HAVE_PODOFO)
                        docx::render_native_pdf(document_, parts_, target);
#endif
                        break;
                    case ConverterKind::Remote:
#if defined(DOCWEFT_HAVE_REMOTE_CONVERTER)
                        error = run_remote_conversion(conv.location, scratch_docx, target);
#endif
                        break;
                }
            } catch (const ReportException& e) {
                code  = e.code();
                error = e.what();
            }
            if (error.empty() && !std::filesystem::exists(target)) {
                error = "reported success but produced no output";
            }
            if (error.empty()) {
                produced = target;
            } else {
                failures.push_back({code, fmt::format("{}: {}",
                                    converter_label(conv.kind), error)});
            }
        }

        if (produced.empty()) {
            cleanup();
            if (failures.size() == 1) {
                throw ReportException(failures.front().code, failures.front().text);
            }
            std::string why;
            for (const auto& f : failures) why += "\n  - " + f.text;
            throw ReportException(
                ReportError::SaveFailed,
                fmt::format("every available converter failed to produce {}:{}",
                            ext, why));
        }

        // LibreOffice's HTML export writes referenced images as sibling
        // files next to `produced` (e.g. "output_html_....png") instead of
        // embedding them. Inline every such file into `produced` as a
        // base64 data: URI so the final HTML is self-contained; the
        // siblings themselves are discarded with the rest of `scratch`
        // below. PDF conversion has no siblings, so this is a no-op for
        // ".pdf". Converter helper files (tf_*) are not images.
        if (conv_format == "html") {
            std::error_code list_ec;
            std::vector<std::filesystem::path> siblings;
            for (const auto& entry :
                 std::filesystem::directory_iterator(produced.parent_path(), list_ec)) {
                if (!entry.is_regular_file()) continue;
                if (entry.path() == produced) continue;
                if (entry.path().filename().string().rfind("tf_", 0) == 0) continue;
                siblings.push_back(entry.path());
            }
            inline_sibling_images(produced, siblings);
        }

        // Move into place. Cross-device move is possible (tmp on a different
        // filesystem than the target) so fall back to copy.
        std::filesystem::remove(p, cleanup_ec);
        std::error_code move_ec;
        std::filesystem::rename(produced, p, move_ec);
        if (move_ec) {
            std::filesystem::copy_file(
                produced, p,
                std::filesystem::copy_options::overwrite_existing, move_ec);
            if (move_ec) {
                cleanup();
                throw ReportException(
                    ReportError::SaveFailed,
                    fmt::format("failed to place converted file at {}: {}",
                                p.string(), move_ec.message()));
            }
        }
        cleanup();
        return;
    }
    if (ext != ".docx") {
        throw ReportException(
            ReportError::SaveFailed,
            fmt::format("unsupported extension: {}", ext));
    }

    reserialize_document();
    write_archive(p);
}

std::string DocxMerger::document_xml() const {
    // Serialize the live DOM — parts_[kDocumentPart] is only refreshed
    // during save(), so it can be stale after mutations.
    if (!loaded_) return {};
    std::ostringstream oss;
    document_.save(oss, "", pugi::format_raw | pugi::format_no_declaration);
    return oss.str();
}

// ── Bookmark scanning ───────────────────────────────────────────────────────

std::vector<pugi::xml_node>
DocxMerger::find_bookmark_starts(const std::string& name) const {
    std::vector<pugi::xml_node> out;
    // Walk the whole document — bookmarks can live anywhere
    // inside <w:body>, tables, headers etc.
    struct Walker : pugi::xml_tree_walker {
        const std::string& target;
        std::vector<pugi::xml_node>& out;
        Walker(const std::string& t, std::vector<pugi::xml_node>& o)
            : target(t), out(o) {}
        bool for_each(pugi::xml_node& n) override {
            if (std::strcmp(n.name(), "w:bookmarkStart") == 0) {
                const auto attr = n.attribute("w:name");
                if (attr && target == attr.value()) out.push_back(n);
            }
            return true;
        }
    } walker{name, out};

    // pugixml's const-walker pattern requires a non-const document reference.
    const_cast<pugi::xml_document&>(document_).traverse(walker);
    return out;
}

// ── setClipboardValue ───────────────────────────────────────────────────────
//
// Strategy for v1:
//   1. Find the <w:bookmarkStart w:name="..."> node matching `bookmark`.
//   2. Walk forward in document order (up to the matching <w:bookmarkEnd>)
//      collecting <w:t> nodes.
//   3. Concatenate their text; replace every occurrence of `field` (not
//      just the first — Word can duplicate the same run inside
//      mc:AlternateContent for old-Word compatibility). Redistribute each
//      match: write its own prefix + value into the node it starts in,
//      clear any nodes it fully consumes, and keep the suffix of the node
//      it ends in. This preserves run-property formatting (the <w:rPr> on
//      the first run of each match).
//
// This is intentionally conservative — it handles both `kFIELD`-style
// and `{{field}}` placeholders inside a single bookmark.

namespace {

// Collect <w:t> nodes between a bookmarkStart and its matching bookmarkEnd.
std::vector<pugi::xml_node>
collect_text_nodes_in_bookmark(pugi::xml_node bookmark_start) {
    std::vector<pugi::xml_node> texts;

    // Walk forward from bookmark_start in document order until we hit
    // <w:bookmarkEnd> with matching w:id.
    const auto id_attr = bookmark_start.attribute("w:id");
    const std::string wanted_id = id_attr ? id_attr.value() : "";

    // Simple depth-first traversal using explicit stack over the subtree
    // that *follows* the bookmark inside its parent chain.
    pugi::xml_node body = bookmark_start;
    while (body && std::strcmp(body.name(), "w:body") != 0 && body.parent()) {
        body = body.parent();
    }
    if (!body) return texts;

    struct Walker : pugi::xml_tree_walker {
        const std::string& wanted_id;
        std::vector<pugi::xml_node>& texts;
        bool started = false;
        bool done    = false;
        pugi::xml_node start_node;

        Walker(const std::string& id, std::vector<pugi::xml_node>& t,
               pugi::xml_node s)
            : wanted_id(id), texts(t), start_node(s) {}

        bool for_each(pugi::xml_node& n) override {
            if (done) return false;
            if (!started) {
                if (n == start_node) started = true;
                return true;
            }
            if (std::strcmp(n.name(), "w:bookmarkEnd") == 0) {
                const auto id = n.attribute("w:id");
                if (id && wanted_id == id.value()) {
                    done = true;
                    return false;
                }
            }
            if (std::strcmp(n.name(), "w:t") == 0) {
                texts.push_back(n);
            }
            return true;
        }
    } walker{wanted_id, texts, bookmark_start};

    body.traverse(walker);
    return texts;
}

// Collect every <w:t> descendant of `root` in document order.
// Used by setTableRow to scope replacements to a single <w:tc>.
std::vector<pugi::xml_node>
collect_text_descendants(pugi::xml_node root) {
    std::vector<pugi::xml_node> out;
    struct Walker : pugi::xml_tree_walker {
        std::vector<pugi::xml_node>& o;
        explicit Walker(std::vector<pugi::xml_node>& o_) : o(o_) {}
        bool for_each(pugi::xml_node& n) override {
            if (std::strcmp(n.name(), "w:t") == 0) o.push_back(n);
            return true;
        }
    } w{out};
    root.traverse(w);
    return out;
}

// Preserve leading / trailing spaces on a <w:t> by setting xml:space="preserve".
// Without this attribute Word collapses boundary whitespace on open.
void ensure_xml_space_preserve(pugi::xml_node t, const std::string& text) {
    if (text.empty()) return;
    if (text.front() != ' ' && text.back() != ' ') return;
    auto attr = t.attribute("xml:space");
    if (!attr) attr = t.append_attribute("xml:space");
    attr.set_value("preserve");
}

// Find every non-overlapping occurrence of `field` in `joined`, left to
// right, skipping any candidate that isn't a real match. A plain substring
// search would match a bare field name (e.g. "kUser") inside the braced
// form of a different placeholder ("{kUser}"), replacing only the letters
// and leaving the original braces stranded around the new value — so skip
// a candidate wrapped in '{'/'}' unless `field` itself already includes
// those braces. Word can also duplicate the same run twice within one part
// (e.g. mc:Choice / mc:Fallback content for old-Word compatibility), so a
// single first-match search misses the second copy entirely — collect all
// of them instead.
std::vector<std::pair<std::size_t, std::size_t>>
find_field_tokens(const std::string& joined, const std::string& field) {
    const bool field_is_braced =
        field.size() >= 2 && field.front() == '{' && field.back() == '}';

    std::vector<std::pair<std::size_t, std::size_t>> matches;
    std::size_t search_from = 0;
    while (true) {
        const auto pos = joined.find(field, search_from);
        if (pos == std::string::npos) break;

        const bool wrapped_in_braces =
            !field_is_braced &&
            pos > 0 && joined[pos - 1] == '{' &&
            pos + field.size() < joined.size() && joined[pos + field.size()] == '}';
        if (wrapped_in_braces) {
            search_from = pos + 1;
            continue;
        }

        matches.emplace_back(pos, pos + field.size());
        search_from = pos + field.size();
    }
    return matches;
}

// Replace every occurrence of `field` across the concatenated text of
// `texts` with `value`. For each match, only the <w:t> nodes it actually
// spans are rewritten: the node the match starts in receives its own
// untouched prefix + `value`, the node it ends in keeps its own untouched
// suffix, and any nodes fully consumed by the match are cleared. Nodes no
// match touches (including ones belonging to unrelated overlapping
// bookmarks) are left byte-for-byte untouched. Returns true when at least
// one replacement happened.
bool replace_placeholder_in(std::vector<pugi::xml_node>& texts,
                            const std::string& field,
                            const std::string& value) {
    if (texts.empty()) return false;

    // Record where each contributor starts in the joined string so we can
    // map matches back to individual nodes later.
    std::string joined;
    std::vector<std::size_t> offsets;
    offsets.reserve(texts.size());
    for (auto t : texts) {
        offsets.push_back(joined.size());
        joined += t.text().get();
    }

    const auto matches = find_field_tokens(joined, field);
    if (matches.empty()) return false;

    auto node_end = [&](std::size_t i) {
        return (i + 1 < offsets.size()) ? offsets[i + 1] : joined.size();
    };

    for (std::size_t i = 0; i < texts.size(); ++i) {
        const std::size_t start_i = offsets[i];
        const std::size_t end_i   = node_end(i);
        const std::string original = texts[i].text().get();

        std::string new_text;
        std::size_t cursor = start_i;
        bool touched = false;

        for (const auto& [m_start, m_end] : matches) {
            if (m_end <= start_i) continue;   // entirely before this node
            if (m_start >= end_i) break;      // entirely after (matches are sorted)
            touched = true;

            const std::size_t seg_start = std::max(m_start, start_i);
            const std::size_t seg_end   = std::min(m_end, end_i);
            new_text += original.substr(cursor - start_i, seg_start - cursor);
            if (m_start >= start_i) new_text += value;  // match starts in this node
            cursor = seg_end;
        }
        if (!touched) continue;

        new_text += original.substr(cursor - start_i);
        texts[i].text().set(new_text.c_str());
        ensure_xml_space_preserve(texts[i], new_text);
    }
    return true;
}

// Walk up the parent chain from `node` until a tag named `tag_name` is found.
// Returns an empty node if no such ancestor exists.
pugi::xml_node ancestor_named(pugi::xml_node node, const char* tag_name) {
    for (auto cur = node; cur; cur = cur.parent()) {
        if (std::strcmp(cur.name(), tag_name) == 0) return cur;
    }
    return pugi::xml_node{};
}

// Delete every <w:bookmarkStart>/<w:bookmarkEnd> descendant of `root`.
// Prevents duplicate `w:id` attributes when a template row is cloned —
// Word otherwise refuses to open the document.
void strip_bookmarks(pugi::xml_node root) {
    std::vector<pugi::xml_node> to_remove;
    struct Walker : pugi::xml_tree_walker {
        std::vector<pugi::xml_node>& kill;
        explicit Walker(std::vector<pugi::xml_node>& k) : kill(k) {}
        bool for_each(pugi::xml_node& n) override {
            const char* nm = n.name();
            if (std::strcmp(nm, "w:bookmarkStart") == 0 ||
                std::strcmp(nm, "w:bookmarkEnd") == 0) {
                kill.push_back(n);
            }
            return true;
        }
    } w{to_remove};
    root.traverse(w);
    for (auto n : to_remove) n.parent().remove_child(n);
}

// Recognize "word/header<N>.xml" / "word/footer<N>.xml" archive parts — the
// only shape Word produces for these header/footer parts, none of which
// carry bookmarks of their own.
bool is_header_footer_part(const std::string& part_name, std::string_view prefix) {
    if (part_name.size() < prefix.size() + 5) return false;   // prefix + digit + ".xml"
    if (part_name.compare(0, prefix.size(), prefix) != 0) return false;
    if (part_name.compare(part_name.size() - 4, 4, ".xml") != 0) return false;
    for (std::size_t i = prefix.size(); i < part_name.size() - 4; ++i) {
        if (part_name[i] < '0' || part_name[i] > '9') return false;
    }
    return true;
}

} // namespace

void DocxMerger::setClipboardValue(const std::string& bookmark,
                                   const std::string& field,
                                   const std::string& value) {
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setClipboardValue before load()");
    }

    // ── "_Header" / "_Footer" pseudo-bookmarks (data contract §3.1) ────────
    // Word keeps header/footer content in separate parts
    // (word/header{N}.xml, word/footer{N}.xml) that document.xml only
    // references from section properties — they typically carry no
    // <w:bookmarkStart> of their own. There's nothing to
    // resolve via find_bookmark_starts, so instead we walk every part whose
    // name matches the requested half and replace `field` in each one that
    // contains it, using the same split-run-safe replacement as everywhere
    // else.
    if (bookmark == "_Header" || bookmark == "_Footer") {
        const std::string_view prefix =
            (bookmark == "_Header") ? std::string_view("word/header")
                                     : std::string_view("word/footer");

        bool any_part      = false;
        bool any_replaced  = false;
        for (auto& [name, bytes] : parts_) {
            if (!is_header_footer_part(name, prefix)) continue;
            any_part = true;

            pugi::xml_document part_doc;
            if (!part_doc.load_buffer(bytes.data(), bytes.size())) continue;

            auto texts = collect_text_descendants(part_doc.document_element());
            if (replace_placeholder_in(texts, field, value)) {
                std::ostringstream oss;
                part_doc.save(oss, "", pugi::format_raw);
                bytes = oss.str();
                any_replaced = true;
            }
        }

        if (!any_part) {
            throw ReportException(
                ReportError::InvalidBookmark,
                fmt::format("no {} parts in archive", bookmark));
        }
        if (!any_replaced) {
            throw ReportException(
                ReportError::InvalidField,
                fmt::format("field '{}' not found in any {} part", field, bookmark));
        }
        return;
    }

    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }

    bool any_replaced = false;

    for (auto bmk : starts) {
        auto texts = collect_text_nodes_in_bookmark(bmk);
        if (replace_placeholder_in(texts, field, value)) {
            any_replaced = true;
        }
    }

    if (!any_replaced) {
        throw ReportException(
            ReportError::InvalidField,
            fmt::format("field '{}' not found in bookmark '{}'", field, bookmark));
    }
}

bool DocxMerger::hasBookmark(const std::string& bookmark) const {
    if (!loaded_) return false;
    return !find_bookmark_starts(bookmark).empty();
}

void DocxMerger::clearBookmark(const std::string& bookmark) {
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "clearBookmark before load()");
    }

    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }

    // Blank every <w:t> the bookmark spans — same span collect_text_nodes_in_bookmark
    // uses for setClipboardValue, so a template that round-trips through
    // setClipboardValue round-trips through clearBookmark too.
    for (auto bmk : starts) {
        for (auto t : collect_text_nodes_in_bookmark(bmk)) {
            t.text().set("");
        }
    }
}

// ── setTableRow ─────────────────────────────────────────────────────────────
//
// Stage 3 — clones a <w:tr> template row once per entry in `rows`,
// substituting `fields[j]` with `rows[i][j]` inside the j-th <w:tc> that
// contains the placeholder.
//   - bookmark must be inside exactly one template <w:tr>
//   - clones go immediately after the template row, in input order
//   - the template row itself is removed after cloning (so rows.empty()
//     yields an empty table body)
//   - cloned rows have their <w:bookmarkStart/End> stripped to keep w:id unique

void DocxMerger::setTableRow(
    const std::string& bookmark,
    const std::vector<std::string>& fields,
    const std::vector<std::vector<std::string>>& rows)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setTableRow before load()");
    }

    // Validate row widths up-front so we never leave the document in a
    // half-mutated state on mismatched input.
    for (std::size_t i = 0; i < rows.size(); ++i) {
        if (rows[i].size() != fields.size()) {
            throw ReportException(
                ReportError::InvalidField,
                fmt::format("setTableRow: rows[{}].size={} does not match fields.size={}",
                            i, rows[i].size(), fields.size()));
        }
    }

    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }

    pugi::xml_node template_tr = ancestor_named(starts.front(), "w:tr");
    if (!template_tr) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark '{}' is not inside a <w:tr> row", bookmark));
    }
    pugi::xml_node tbl = template_tr.parent();
    if (!tbl || std::strcmp(tbl.name(), "w:tbl") != 0) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("row for bookmark '{}' has no <w:tbl> parent", bookmark));
    }

    pugi::xml_node anchor = template_tr;
    for (std::size_t i = 0; i < rows.size(); ++i) {
        pugi::xml_node clone = tbl.insert_copy_after(template_tr, anchor);
        strip_bookmarks(clone);

        // Replace each field inside the first <w:tc> that still contains it.
        // Per-cell scoping is required because replace_placeholder_in merges
        // all text into the first <w:t> of its input; spanning cells would
        // destroy their content.
        for (std::size_t j = 0; j < fields.size(); ++j) {
            bool replaced = false;
            for (auto tc = clone.child("w:tc"); tc;
                 tc = tc.next_sibling("w:tc")) {
                auto texts = collect_text_descendants(tc);
                if (replace_placeholder_in(texts, fields[j], rows[i][j])) {
                    replaced = true;
                    break;
                }
            }
            if (!replaced) {
                throw ReportException(
                    ReportError::InvalidField,
                    fmt::format("field '{}' not found in any cell of row template",
                                fields[j]));
            }
        }
        anchor = clone;
    }

    tbl.remove_child(template_tr);
}

// ── setImage (Stage 5) ──────────────────────────────────────────────────────
//
// Inserts a <w:drawing> into the paragraph holding `bookmark`, dropping any
// text children the bookmark used to contain (so a template placeholder like
// "{{image}}" is fully replaced). Adds the PNG bytes as a new archive part
// `word/media/imageN.png`, registers a relationship in
// `word/_rels/document.xml.rels`, and ensures `[Content_Types].xml`
// advertises `png`.

namespace {

constexpr const char* kDocRelsPart     = "word/_rels/document.xml.rels";
constexpr const char* kContentTypesPart = "[Content_Types].xml";
constexpr const char* kImageRelType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

// 914400 EMU/inch ÷ 96 px/inch. Word renders `<wp:extent>` at this scale
// by default — image appears at 1:1 with pixel size on a 96-DPI display.
constexpr std::uint32_t kEmuPerPx = 9525;

// Pick the next free `word/media/image{N}.<ext>` that doesn't collide with an
// existing archive part.
std::string next_media_name(
    const std::unordered_map<std::string, std::string>& parts,
    const std::string& ext /* "png" */)
{
    for (int n = 1; n < 10000; ++n) {
        auto candidate = fmt::format("word/media/image{}.{}", n, ext);
        if (parts.find(candidate) == parts.end()) return candidate;
    }
    throw ReportException(ReportError::SaveFailed,
                          "unable to allocate a free image part name");
}

// Parse existing rels file to find the highest rId numeric suffix so new
// relationships can safely append. Returns 0 when no numeric rIds are present.
int max_numeric_rid(pugi::xml_document& rels) {
    int m = 0;
    for (auto rel : rels.child("Relationships").children("Relationship")) {
        const auto id = rel.attribute("Id");
        if (!id) continue;
        const std::string v = id.value();
        if (v.size() > 3 && v.compare(0, 3, "rId") == 0) {
            try {
                const int n = std::stoi(v.substr(3));
                if (n > m) m = n;
            } catch (...) { /* non-numeric, skip */ }
        }
    }
    return m;
}

// Make sure `word/_rels/document.xml.rels` exists in the archive. Creates a
// minimal one when the template didn't ship any relationships.
void ensure_doc_rels_present(
    std::unordered_map<std::string, std::string>& parts)
{
    if (parts.find(kDocRelsPart) != parts.end()) return;
    parts[kDocRelsPart] =
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>)";
}

// Register a new <Relationship Id=... Type=... Target=.../> and return the id.
std::string add_image_relationship(
    std::unordered_map<std::string, std::string>& parts,
    const std::string& target_inside_word /* e.g. "media/image1.png" */)
{
    ensure_doc_rels_present(parts);
    pugi::xml_document rels;
    const auto& xml = parts[kDocRelsPart];
    if (!rels.load_buffer(xml.data(), xml.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "document.xml.rels is not valid XML");
    }
    auto root = rels.child("Relationships");
    if (!root) {
        root = rels.append_child("Relationships");
        root.append_attribute("xmlns").set_value(
            "http://schemas.openxmlformats.org/package/2006/relationships");
    }
    const int next_id = max_numeric_rid(rels) + 1;
    const std::string rid = fmt::format("rId{}", next_id);
    auto r = root.append_child("Relationship");
    r.append_attribute("Id").set_value(rid.c_str());
    r.append_attribute("Type").set_value(kImageRelType);
    r.append_attribute("Target").set_value(target_inside_word.c_str());

    std::ostringstream oss;
    rels.save(oss, "", pugi::format_raw);
    parts[kDocRelsPart] = oss.str();
    return rid;
}

// Ensure [Content_Types].xml has <Default Extension="png" ContentType="image/png"/>.
// No-op when the entry is already present; appends otherwise.
void ensure_content_type(
    std::unordered_map<std::string, std::string>& parts,
    const std::string& extension     /* "png" */,
    const std::string& content_type  /* "image/png" */)
{
    auto it = parts.find(kContentTypesPart);
    if (it == parts.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "[Content_Types].xml is missing from the archive");
    }
    pugi::xml_document ct;
    if (!ct.load_buffer(it->second.data(), it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "[Content_Types].xml is not valid XML");
    }
    auto types = ct.child("Types");
    if (!types) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "[Content_Types].xml has no <Types> root");
    }
    for (auto def : types.children("Default")) {
        const auto ext = def.attribute("Extension");
        if (ext && extension == ext.value()) return;  // already declared
    }
    auto def = types.append_child("Default");
    def.append_attribute("Extension").set_value(extension.c_str());
    def.append_attribute("ContentType").set_value(content_type.c_str());

    std::ostringstream oss;
    ct.save(oss, "", pugi::format_raw);
    it->second = oss.str();
}

// Compute a single uniform scale factor that maps (native_w, native_h)
// inside the user's ImageSize bounds, preserving aspect ratio.
//
// The problem reduces to a feasible range [s_lo, s_hi] where:
//   s_hi = min(max_w / w, max_h / h)   — tightest max constraint (∞ if none)
//   s_lo = max(min_w / w, min_h / h)   — tightest min constraint (0 if none)
//
// If s_lo > s_hi, the bounds contradict through the aspect ratio (e.g. user
// asked for `min_width_px=400, max_height_px=200` on a square image).
// We surface that as InvalidField rather than silently picking a scale.
// Otherwise we clamp the intrinsic 1.0 scale into [s_lo, s_hi].
double compute_display_scale(std::uint32_t w, std::uint32_t h,
                             const ImageSize& b)
{
    if (w == 0 || h == 0) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "image has zero dimensions");
    }

    double s_hi = std::numeric_limits<double>::infinity();
    if (b.max_width_px  > 0) s_hi = std::min(s_hi, static_cast<double>(b.max_width_px)  / w);
    if (b.max_height_px > 0) s_hi = std::min(s_hi, static_cast<double>(b.max_height_px) / h);

    double s_lo = 0.0;
    if (b.min_width_px  > 0) s_lo = std::max(s_lo, static_cast<double>(b.min_width_px)  / w);
    if (b.min_height_px > 0) s_lo = std::max(s_lo, static_cast<double>(b.min_height_px) / h);

    if (s_lo > s_hi) {
        throw ReportException(
            ReportError::InvalidField,
            fmt::format("ImageSize bounds have empty feasible range: "
                        "min scale {:.4f} > max scale {:.4f} for a "
                        "{}×{} image", s_lo, s_hi, w, h));
    }

    // Intrinsic scale = 1.0. Keep it if it's inside the feasible band,
    // otherwise clamp to the nearest border (the one the user actually hit).
    if (1.0 < s_lo) return s_lo;
    if (1.0 > s_hi) return s_hi;
    return 1.0;
}

// Build the inline <w:drawing> XML fragment that embeds a picture.
// Standard OOXML DrawingML boilerplate — the parameters are the only moving
// parts: unique doc-pr id, relationship id, and EMU-scale extent.
std::string make_drawing_xml(std::uint32_t uid,
                             const std::string& rid,
                             std::uint32_t cx_emu,
                             std::uint32_t cy_emu)
{
    return fmt::format(
R"(<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{id}" name="Picture {id}"/>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr>
            <pic:cNvPr id="{id}" name="Picture {id}"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="{cx}" cy="{cy}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>)",
        fmt::arg("cx", cx_emu),
        fmt::arg("cy", cy_emu),
        fmt::arg("id", uid),
        fmt::arg("rid", rid));
}

} // namespace

void DocxMerger::setImage(const std::string& bookmark,
                          const std::filesystem::path& image_path,
                          const ImageSize& bounds)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setImage before load()");
    }

    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }
    pugi::xml_node bmk_start = starts.front();
    pugi::xml_node para = ancestor_named(bmk_start, "w:p");
    if (!para) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark '{}' is not inside a <w:p> paragraph — "
                        "setImage requires a paragraph-scoped bookmark",
                        bookmark));
    }

    // Read + normalise the source image (PNG pass-through in Phase 1).
    const PngBuffer img = load_as_png(image_path);

    // Resolve bounds → single aspect-preserving scale factor. Must happen
    // BEFORE any archive mutation so an invalid ImageSize doesn't leave the
    // document with a dangling media part / relationship.
    const double scale = compute_display_scale(img.width, img.height, bounds);

    // Commit to the archive: pick media path, register relationship, ensure
    // content-type. Any throw past this point leaves the document untouched
    // from the caller's perspective (the archive mutations happen via parts_
    // which is only written back on save()).
    const std::string media_part = next_media_name(parts_, "png");
    parts_[media_part] = img.bytes;

    // Target inside a Word relationship is relative to the `word/` folder,
    // so strip the "word/" prefix when referencing the media.
    const std::string target_inside_word = media_part.substr(std::strlen("word/"));
    const std::string rid = add_image_relationship(parts_, target_inside_word);
    ensure_content_type(parts_, "png", "image/png");

    // Compute EMU extents (9525 EMU/px at 96 DPI) from the scaled pixel size.
    // Cap at a sane maximum so a pathological 50000×50000 image doesn't
    // produce values Word rejects.
    constexpr std::uint64_t kMaxEmu = 360'000'000ULL;  // ~394 inches
    const double scaled_w = static_cast<double>(img.width)  * scale;
    const double scaled_h = static_cast<double>(img.height) * scale;
    auto cx = static_cast<std::uint64_t>(std::llround(scaled_w * kEmuPerPx));
    auto cy = static_cast<std::uint64_t>(std::llround(scaled_h * kEmuPerPx));
    if (cx == 0) cx = kEmuPerPx;          // never emit a zero-size extent
    if (cy == 0) cy = kEmuPerPx;
    if (cx > kMaxEmu) cx = kMaxEmu;
    if (cy > kMaxEmu) cy = kMaxEmu;

    // Build the <w:drawing> once and parse it into the live document. A
    // counter-based uid prevents collisions with previously-inserted pictures.
    static std::uint32_t docPrCounter = 1000;
    const std::uint32_t uid = ++docPrCounter;
    const std::string draw_xml = make_drawing_xml(
        uid, rid, static_cast<std::uint32_t>(cx), static_cast<std::uint32_t>(cy));

    pugi::xml_document drawing_doc;
    // The fragment uses `w:`, `wp:`, `a:`, `pic:`, `r:` prefixes. These
    // namespaces are already declared on the root <w:document> element, so
    // parsing as a fragment with default options is safe.
    const auto parse_result = drawing_doc.load_buffer(
        draw_xml.data(), draw_xml.size(),
        pugi::parse_default | pugi::parse_ws_pcdata_single);
    if (!parse_result) {
        throw ReportException(
            ReportError::SaveFailed,
            fmt::format("internal: failed to parse drawing XML: {}",
                        parse_result.description()));
    }

    // Strip all <w:r> inside the paragraph whose bookmark we hit. This covers
    // the common case of a placeholder like "{{image}}" inside the bookmark —
    // the text runs are replaced in-full by the picture.
    std::vector<pugi::xml_node> to_remove;
    for (auto child = para.first_child(); child; child = child.next_sibling()) {
        const char* nm = child.name();
        if (std::strcmp(nm, "w:r") == 0) to_remove.push_back(child);
    }
    for (auto n : to_remove) para.remove_child(n);

    // Insert a fresh <w:r> containing the <w:drawing> right before the
    // bookmarkEnd (fallback: after bookmarkStart) so it lands between the
    // bookmark markers.
    pugi::xml_node bmk_end;
    const auto bmk_id = bmk_start.attribute("w:id");
    if (bmk_id) {
        for (auto sib = bmk_start; sib; sib = sib.next_sibling()) {
            if (std::strcmp(sib.name(), "w:bookmarkEnd") == 0) {
                const auto id = sib.attribute("w:id");
                if (id && std::strcmp(id.value(), bmk_id.value()) == 0) {
                    bmk_end = sib;
                    break;
                }
            }
        }
    }
    pugi::xml_node run = bmk_end
        ? para.insert_child_before("w:r", bmk_end)
        : para.insert_child_after("w:r", bmk_start);
    run.append_copy(drawing_doc.first_child());
}

// ── setChartValue (Stage 4) ─────────────────────────────────────────────────
//
// Updates a single numeric point inside an embedded chart. The flow:
//
//   bookmark → <w:p> ancestor → descendant <c:chart r:id="rIdX"/>
//            → resolve rIdX via word/_rels/document.xml.rels → chart part
//            → parse XML → match (series, category) in any series container
//              that uses the <c:cat>/<c:val> shape (bar/line/pie/area/…)
//            → rewrite <c:val>/<c:numRef>/<c:numCache>/<c:pt idx=N>/<c:v>
//
// The `field` parameter is accepted for API symmetry with the other setters
// but does not participate in the match — callers typically pass a literal
// constant that never varies per call, so using it for selection would just
// couple implementation to templates without adding discriminating power.
//
// Embedded spreadsheets (word/embeddings/*.xlsx) are not touched. Word may
// rebuild the chart cache from them when the template is manually edited;
// report templates should keep the embedding read-only or omit it.

namespace {

// Look up the Target of the Relationship with matching Id in the given rels
// XML blob. Returns the empty string when the rId isn't present.
std::string resolve_doc_rid(const std::string& rels_xml,
                            const std::string& rid)
{
    pugi::xml_document rels;
    if (!rels.load_buffer(rels_xml.data(), rels_xml.size())) return {};
    for (auto rel : rels.child("Relationships").children("Relationship")) {
        if (rid == rel.attribute("Id").value()) {
            return rel.attribute("Target").value();
        }
    }
    return {};
}

// Find the first <c:chart> node reachable from `root` (inclusive). pugi is
// xmlns-unaware so we compare the literal prefixed tag name.
pugi::xml_node find_descendant_chart(pugi::xml_node root) {
    struct Finder : pugi::xml_tree_walker {
        pugi::xml_node hit;
        bool for_each(pugi::xml_node& n) override {
            if (std::strcmp(n.name(), "c:chart") == 0) {
                hit = n;
                return false;
            }
            return true;
        }
    } f;
    root.traverse(f);
    return f.hit;
}

// Read <c:pt idx="0"><c:v>NAME</c:v></c:pt> — the canonical layout for a
// series display name under <c:tx>/<c:strRef>/<c:strCache>.
std::string read_series_name(pugi::xml_node ser) {
    auto pt = ser.child("c:tx").child("c:strRef").child("c:strCache").child("c:pt");
    if (!pt) {
        // Rare: literal <c:tx><c:v>Name</c:v></c:tx> form.
        return ser.child("c:tx").child_value("c:v");
    }
    return pt.child_value("c:v");
}

// Locate the <c:pt idx="I"> whose <c:v> text equals `category` inside the
// series' <c:cat> block. Categories may be string- or numeric-typed; we try
// both. Returns -1 when no match.
int find_category_idx(pugi::xml_node ser, const std::string& category) {
    auto cat = ser.child("c:cat");
    if (!cat) return -1;
    auto cache = cat.child("c:strRef").child("c:strCache");
    if (!cache) cache = cat.child("c:numRef").child("c:numCache");
    if (!cache) return -1;
    for (auto pt : cache.children("c:pt")) {
        if (category == pt.child_value("c:v")) {
            return pt.attribute("idx").as_int(-1);
        }
    }
    return -1;
}

// Rewrite the numeric cache point at `cat_idx` for `value`. Throws
// CantCopyDocxTemplate if the cache shape is unexpected.
void write_value_point(pugi::xml_node ser, int cat_idx, double value) {
    auto cache = ser.child("c:val").child("c:numRef").child("c:numCache");
    if (!cache) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "series is missing <c:val>/<c:numRef>/<c:numCache>");
    }
    std::ostringstream ss;
    ss.imbue(std::locale::classic());
    ss << value;
    const std::string formatted = ss.str();
    for (auto pt : cache.children("c:pt")) {
        if (pt.attribute("idx").as_int(-1) == cat_idx) {
            if (auto v = pt.child("c:v")) {
                v.text().set(formatted.c_str());
            } else {
                pt.append_child("c:v").text().set(formatted.c_str());
            }
            return;
        }
    }
    throw ReportException(
        ReportError::CantCopyDocxTemplate,
        fmt::format("no <c:val>/<c:pt idx='{}'> to update", cat_idx));
}

bool is_supported_chart_type(std::string_view tag) {
    return tag == "c:barChart"      || tag == "c:lineChart"
        || tag == "c:pieChart"      || tag == "c:areaChart"
        || tag == "c:doughnutChart" || tag == "c:radarChart"
        || tag == "c:ofPieChart"
        || tag == "c:bar3DChart"    || tag == "c:line3DChart"
        || tag == "c:pie3DChart"    || tag == "c:area3DChart";
}

bool is_unsupported_xy_chart(std::string_view tag) {
    return tag == "c:scatterChart" || tag == "c:bubbleChart"
        || tag == "c:stockChart"
        || tag == "c:surfaceChart" || tag == "c:surface3DChart";
}

} // namespace

std::string DocxMerger::locate_chart_part(const std::string& bookmark) const {
    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }
    pugi::xml_node para = ancestor_named(starts.front(), "w:p");
    if (!para) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark '{}' is not inside a <w:p> paragraph",
                        bookmark));
    }

    pugi::xml_node chart_ref = find_descendant_chart(para);
    if (!chart_ref) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark '{}' has no <c:chart> inside its paragraph",
                        bookmark));
    }
    const std::string rid = chart_ref.attribute("r:id").value();
    if (rid.empty()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "<c:chart> element missing r:id attribute");
    }

    auto rels_it = parts_.find(kDocRelsPart);
    if (rels_it == parts_.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "document.xml.rels missing — cannot resolve chart rId");
    }
    const std::string target = resolve_doc_rid(rels_it->second, rid);
    if (target.empty()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("relationship '{}' not declared in document.xml.rels", rid));
    }

    // Relationship Target is relative to word/ — e.g. "charts/chart1.xml".
    const std::string chart_part = "word/" + target;
    if (parts_.find(chart_part) == parts_.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part not in archive: {}", chart_part));
    }
    return chart_part;
}

void DocxMerger::setChartValue(const std::string& bookmark,
                               const std::string& /*field*/,
                               const std::string& series,
                               const std::string& category,
                               double             value)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setChartValue before load()");
    }

    const std::string chart_part = locate_chart_part(bookmark);
    auto chart_it = parts_.find(chart_part);

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(chart_it->second.data(), chart_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part is not valid XML: {}", chart_part));
    }
    auto plot_area = chart_doc.child("c:chartSpace").child("c:chart").child("c:plotArea");
    if (!plot_area) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no <c:plotArea>");
    }

    bool saw_unsupported = false;
    for (auto type_node : plot_area.children()) {
        const std::string_view tag = type_node.name();
        if (is_unsupported_xy_chart(tag)) {
            saw_unsupported = true;
            continue;
        }
        if (!is_supported_chart_type(tag)) continue;

        for (auto ser : type_node.children("c:ser")) {
            if (read_series_name(ser) != series) continue;

            const int cat_idx = find_category_idx(ser, category);
            if (cat_idx < 0) {
                throw ReportException(
                    ReportError::InvalidField,
                    fmt::format("category '{}' not found in series '{}'",
                                category, series));
            }
            write_value_point(ser, cat_idx, value);

            std::ostringstream oss;
            chart_doc.save(oss, "", pugi::format_raw);
            parts_[chart_part] = oss.str();
            return;
        }
    }

    if (saw_unsupported) {
        throw ReportException(
            ReportError::NotImplemented,
            "chart uses scatter/bubble/stock/surface plot — setChartValue only "
            "supports cat/val shapes (bar/line/pie/area/radar/doughnut/3D variants)");
    }
    throw ReportException(
        ReportError::InvalidField,
        fmt::format("series '{}' not found in chart '{}'", series, chart_part));
}

// ── setChartSeriesName / setChartTitle / setChartAxisTitle / setChartData ───
//
// Shared helpers. setChartSeriesName rewrites the same <c:tx>/<c:strRef>/
// <c:strCache> shape write_value_point's sibling already reads via
// read_series_name(). setChartTitle/setChartAxisTitle build a plain
// single-run DrawingML <c:title>. setChartData fully replaces the category
// axis and series set — unlike setChartValue, it can change point counts —
// and, when the chart carries an embedded workbook (word/embeddings/*.xlsx,
// linked through <c:externalData>), rewrites that workbook's backing sheet
// too, via a temp-file round trip through libzip (the simplest way to get
// at a nested zip archive with the same read/write primitives read_archive/
// write_archive already use for the outer .docx).

namespace {

// 0 → "A", 1 → "B", ..., 25 → "Z", 26 → "AA", ... — spreadsheet column
// letters for a 0-based column index.
std::string column_letter(std::size_t zero_based_index) {
    std::string s;
    long n = static_cast<long>(zero_based_index) + 1;
    while (n > 0) {
        const long rem = (n - 1) % 26;
        s.insert(s.begin(), static_cast<char>('A' + rem));
        n = (n - 1) / 26;
    }
    return s;
}

// Extracts the sheet name from a chart formula like "Sheet1!$B$2:$B$11" or
// "'My Sheet'!$B$2:$B$11" (quoted form, with '' as an escaped quote).
// Returns "" if `f` has no '!' separator.
std::string sheet_name_from_formula(const std::string& f) {
    const auto bang = f.find('!');
    if (bang == std::string::npos) return {};
    std::string name = f.substr(0, bang);
    if (name.size() >= 2 && name.front() == '\'' && name.back() == '\'') {
        name = name.substr(1, name.size() - 2);
        std::string out;
        for (std::size_t i = 0; i < name.size(); ++i) {
            if (name[i] == '\'' && i + 1 < name.size() && name[i + 1] == '\'') {
                out += '\'';
                ++i;
            } else {
                out += name[i];
            }
        }
        return out;
    }
    return name;
}

// Rewrite a series' display name — same <c:tx>/<c:strRef>/<c:strCache>/
// <c:pt>/<c:v> shape read_series_name() reads. Builds the cache from
// scratch if the series had none at all (rare/malformed template).
void write_series_name(pugi::xml_node ser, const std::string& new_name) {
    auto tx = ser.child("c:tx");
    if (!tx) tx = ser.prepend_child("c:tx");
    auto ref = tx.child("c:strRef");
    if (!ref) ref = tx.append_child("c:strRef");
    auto cache = ref.child("c:strCache");
    if (!cache) cache = ref.append_child("c:strCache");
    if (!cache.child("c:ptCount")) {
        cache.prepend_child("c:ptCount").append_attribute("val") = 1;
    }
    auto pt = cache.child("c:pt");
    if (!pt) {
        pt = cache.append_child("c:pt");
        pt.append_attribute("idx") = 0;
    }
    if (auto v = pt.child("c:v")) {
        v.text().set(new_name.c_str());
    } else {
        pt.append_child("c:v").text().set(new_name.c_str());
    }
}

// Fully replace a <c:cat> (categories) cache with `values` — new <c:f>
// range and a freshly built <c:strCache> with exactly values.size() points.
// Unlike write_value_point (single-point in-place update), this can change
// the point count.
void rewrite_string_cache(pugi::xml_node cat, const std::string& range,
                          const std::vector<std::string>& values) {
    if (!cat) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "series is missing <c:cat>");
    }
    auto ref = cat.child("c:strRef");
    if (!ref) {
        // A purely-numeric category axis uses <c:numRef> instead; switch to
        // <c:strRef> since setChartData writes text labels.
        if (auto num_ref = cat.child("c:numRef")) cat.remove_child(num_ref);
        ref = cat.prepend_child("c:strRef");
    }
    if (auto f = ref.child("c:f")) {
        f.text().set(range.c_str());
    } else {
        ref.prepend_child("c:f").text().set(range.c_str());
    }

    if (auto old_cache = ref.child("c:strCache")) ref.remove_child(old_cache);
    auto cache = ref.append_child("c:strCache");
    cache.append_child("c:ptCount").append_attribute("val") =
        static_cast<unsigned>(values.size());
    for (std::size_t i = 0; i < values.size(); ++i) {
        auto pt = cache.append_child("c:pt");
        pt.append_attribute("idx") = static_cast<unsigned>(i);
        pt.append_child("c:v").text().set(values[i].c_str());
    }
}

// Fully replace a <c:val> numeric cache the same way, preserving the
// original <c:formatCode> (defaulting to "General" if there wasn't one).
void rewrite_numeric_cache(pugi::xml_node val, const std::string& range,
                           const std::vector<double>& values) {
    if (!val) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "series is missing <c:val>");
    }
    auto ref = val.child("c:numRef");
    if (!ref) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "series' <c:val> is missing <c:numRef>");
    }
    if (auto f = ref.child("c:f")) {
        f.text().set(range.c_str());
    } else {
        ref.prepend_child("c:f").text().set(range.c_str());
    }

    std::string format_code = "General";
    if (auto old_cache = ref.child("c:numCache")) {
        if (auto fc = old_cache.child("c:formatCode")) format_code = fc.child_value();
        ref.remove_child(old_cache);
    }
    auto cache = ref.append_child("c:numCache");
    cache.append_child("c:formatCode").text().set(format_code.c_str());
    cache.append_child("c:ptCount").append_attribute("val") =
        static_cast<unsigned>(values.size());
    for (std::size_t i = 0; i < values.size(); ++i) {
        auto pt = cache.append_child("c:pt");
        pt.append_attribute("idx") = static_cast<unsigned>(i);
        std::ostringstream ss;
        ss.imbue(std::locale::classic());
        ss << values[i];
        pt.append_child("c:v").text().set(ss.str().c_str());
    }
}

// Replace `title`'s <c:tx> with a single-run DrawingML rich-text block
// containing `text`. This is a full text overwrite, not a find/replace
// inside existing runs — a template title with several differently-
// formatted runs ends up as one run either way, since we're setting one
// plain string. But the box-level formatting (<a:bodyPr> — rotation,
// autofit, ...; <a:lstStyle>), the paragraph's <a:pPr> (alignment, ...) and
// the *first* run's <a:rPr> (font/size/color/bold/...) are carried over
// from whatever <c:tx>/<c:rich> already existed, so a template's chosen
// title styling survives a programmatic setChartTitle/setChartAxisTitle
// call instead of being reset to bare defaults. A brand-new title (no
// prior <c:tx>) gets the minimal empty skeleton the schema requires.
void set_rich_title_text(pugi::xml_node title, const std::string& text) {
    auto old_tx   = title.child("c:tx");
    auto old_rich = old_tx ? old_tx.child("c:rich") : pugi::xml_node{};

    pugi::xml_node old_body_pr, old_lst_style, old_p_pr, old_run_pr;
    if (old_rich) {
        old_body_pr   = old_rich.child("a:bodyPr");
        old_lst_style = old_rich.child("a:lstStyle");
        if (auto old_p = old_rich.child("a:p")) {
            old_p_pr = old_p.child("a:pPr");
            if (auto old_r = old_p.child("a:r")) old_run_pr = old_r.child("a:rPr");
        }
    }

    auto tx   = title.prepend_child("c:tx");
    auto rich = tx.append_child("c:rich");
    if (old_body_pr)   rich.append_copy(old_body_pr);
    else               rich.append_child("a:bodyPr");
    if (old_lst_style) rich.append_copy(old_lst_style);
    else               rich.append_child("a:lstStyle");
    auto p = rich.append_child("a:p");
    if (old_p_pr) p.append_copy(old_p_pr);
    auto r = p.append_child("a:r");
    if (old_run_pr) r.append_copy(old_run_pr);
    r.append_child("a:t").text().set(text.c_str());

    if (old_tx) title.remove_child(old_tx);
}

// First child of `ax` (a <c:catAx>/<c:valAx>) that a new <c:title> must be
// inserted before, per CT_CatAx/CT_ValAx's element sequence. <c:crossAx> is
// required by the schema, so it's always a valid fallback anchor.
pugi::xml_node axis_title_anchor(pugi::xml_node ax) {
    for (const char* name : {"c:majorGridlines", "c:minorGridlines", "c:numFmt",
                             "c:majorTickMark", "c:minorTickMark", "c:tickLblPos",
                             "c:spPr", "c:txPr", "c:crossAx"}) {
        if (auto n = ax.child(name)) return n;
    }
    return {};
}

// Set (creating if absent) the title of `chart_node`'s <c:title> itself, or
// an axis's <c:title> when `ax` is non-null — shared by setChartTitle and
// setChartAxisTitle. `insert_before` is where a brand-new <c:title> goes
// (schema-position anchor); for the chart's own title this also clears
// <c:autoTitleDeleted>.
void set_title(pugi::xml_node owner, pugi::xml_node insert_before, const std::string& text) {
    auto title_node = owner.child("c:title");
    if (!title_node) {
        if (!insert_before) {
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  "chart is missing a required element to anchor a new <c:title> at");
        }
        title_node = owner.insert_child_before("c:title", insert_before);
    }
    set_rich_title_text(title_node, text);
}

// True if `plot_area` contains at least one supported (cat/val-shaped)
// chart-type group, out to `saw_unsupported` whether it also/instead
// contains a scatter/bubble/stock/surface group — shared NotImplemented
// gate for every setChart*() entry point that doesn't need a specific
// <c:ser> (title/axis-title setters just need to know the chart is a
// reshape-able kind at all).
bool has_supported_chart_type(pugi::xml_node plot_area, bool& saw_unsupported) {
    saw_unsupported = false;
    bool saw_supported = false;
    for (auto type_node : plot_area.children()) {
        const std::string_view tag = type_node.name();
        if (is_unsupported_xy_chart(tag)) saw_unsupported = true;
        else if (is_supported_chart_type(tag)) saw_supported = true;
    }
    return saw_supported;
}

/// RAII temp file — used to round-trip an in-memory zip blob (an embedded
/// .xlsx living inside the outer .docx zip) through libzip's path-based
/// API, since that's the same read/write primitive read_archive/
/// write_archive already use and avoids pulling in libzip's separate
/// in-memory zip_source API for what's a rare, small nested archive.
class TempFile {
public:
    explicit TempFile(const char* suffix) {
        const auto unique = std::chrono::steady_clock::now().time_since_epoch().count();
        path_ = std::filesystem::temp_directory_path() /
                fmt::format("docweft_{}_{}{}", unique,
                           reinterpret_cast<std::uintptr_t>(this), suffix);
    }
    ~TempFile() {
        std::error_code ec;
        std::filesystem::remove(path_, ec);
    }
    TempFile(const TempFile&) = delete;
    TempFile& operator=(const TempFile&) = delete;

    [[nodiscard]] const std::filesystem::path& path() const noexcept { return path_; }

private:
    std::filesystem::path path_;
};

std::unordered_map<std::string, std::string> read_nested_zip(const std::string& bytes) {
    TempFile tmp(".xlsx");
    {
        std::ofstream f(tmp.path(), std::ios::binary);
        f.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
    }
    int err = 0;
    zip_t* raw = zip_open(tmp.path().string().c_str(), ZIP_RDONLY, &err);
    if (!raw) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "embedded workbook is not a valid .xlsx (zip_open failed)");
    }
    ZipFile zf{raw};
    return slurp_archive(zf.get());
}

std::string write_nested_zip(const std::unordered_map<std::string, std::string>& parts) {
    TempFile tmp(".xlsx");
    int err = 0;
    zip_t* raw = zip_open(tmp.path().string().c_str(), ZIP_CREATE | ZIP_EXCL, &err);
    if (!raw) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "failed to create embedded workbook archive");
    }
    for (const auto& [name, data] : parts) {
        zip_source_t* src = zip_source_buffer(raw, data.data(), data.size(), 0);
        if (!src) {
            zip_discard(raw);
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  fmt::format("zip_source_buffer failed for {}", name));
        }
        if (zip_file_add(raw, name.c_str(), src, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
            zip_source_free(src);
            zip_discard(raw);
            throw ReportException(ReportError::CantCopyDocxTemplate,
                                  fmt::format("zip_file_add({}) failed for embedded workbook", name));
        }
    }
    if (zip_close(raw) != 0) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "zip_close failed for embedded workbook");
    }
    std::ifstream f(tmp.path(), std::ios::binary);
    std::ostringstream oss;
    oss << f.rdbuf();
    return oss.str();
}

// Archive part name of the .xlsx linked from `chart_part` via
// <c:externalData r:id="..."/>, resolved through the chart's own
// "_rels/<basename>.rels" sibling. Returns "" when the chart has no
// external data at all (a template with no embedded workbook — the common
// case, and not an error).
std::string resolve_chart_external_data(
    const std::unordered_map<std::string, std::string>& parts,
    pugi::xml_node chart_space,
    const std::string& chart_part)
{
    auto ext = chart_space.child("c:externalData");
    if (!ext) return {};
    const std::string rid = ext.attribute("r:id").value();
    if (rid.empty()) return {};

    const auto slash = chart_part.find_last_of('/');
    const std::string dir  = (slash == std::string::npos) ? "" : chart_part.substr(0, slash + 1);
    const std::string base = (slash == std::string::npos) ? chart_part : chart_part.substr(slash + 1);
    const std::string rels_part = dir + "_rels/" + base + ".rels";

    auto rels_it = parts.find(rels_part);
    if (rels_it == parts.end()) return {};
    const std::string target = resolve_doc_rid(rels_it->second, rid);
    if (target.empty()) return {};

    return (std::filesystem::path(dir) / target).lexically_normal().generic_string();
}

// Within an already-unzipped embedded workbook, resolve the worksheet part
// backing `sheet_name` (as declared in xl/workbook.xml). Falls back to the
// first xl/worksheets/*.xml part found when the name can't be resolved —
// an embedded chart workbook is effectively always single-sheet, so this
// is a safety net, not the primary path.
std::string resolve_workbook_sheet_part(
    const std::unordered_map<std::string, std::string>& wb_parts,
    const std::string& sheet_name)
{
    auto wb_it   = wb_parts.find("xl/workbook.xml");
    auto rels_it = wb_parts.find("xl/_rels/workbook.xml.rels");
    if (wb_it != wb_parts.end() && rels_it != wb_parts.end() && !sheet_name.empty()) {
        pugi::xml_document wb_doc;
        if (wb_doc.load_buffer(wb_it->second.data(), wb_it->second.size())) {
            for (auto sheet : wb_doc.child("workbook").child("sheets").children("sheet")) {
                if (sheet_name != sheet.attribute("name").value()) continue;
                const std::string rid = sheet.attribute("r:id").value();
                if (rid.empty()) continue;
                const std::string target = resolve_doc_rid(rels_it->second, rid);
                if (target.empty()) continue;
                return (std::filesystem::path("xl") / target).lexically_normal().generic_string();
            }
        }
    }
    std::vector<std::string> candidates;
    for (const auto& [name, _] : wb_parts) {
        if (name.rfind("xl/worksheets/", 0) == 0 &&
            name.size() > 4 && name.compare(name.size() - 4, 4, ".xml") == 0) {
            candidates.push_back(name);
        }
    }
    if (candidates.empty()) return {};
    std::sort(candidates.begin(), candidates.end());
    return candidates.front();
}

// Full worksheet XML for a categories × series grid, using inline strings
// (<c t="inlineStr">) for every text cell so there's no shared-strings
// table (xl/sharedStrings.xml) to keep in sync on top of the sheet itself.
// Column A holds categories; B, C, ... hold one series each, in order. Row
// 1 is the series-name header row that the chart's <c:tx> ranges point at.
std::string build_worksheet_xml(
    const std::vector<std::string>& categories,
    const std::vector<IReportMerger::ChartSeries>& series)
{
    const std::string dim_ref =
        fmt::format("A1:{}{}", column_letter(series.size()), categories.size() + 1);

    pugi::xml_document doc;
    auto decl = doc.append_child(pugi::node_declaration);
    decl.append_attribute("version")    = "1.0";
    decl.append_attribute("encoding")   = "UTF-8";
    decl.append_attribute("standalone") = "yes";

    auto ws = doc.append_child("worksheet");
    ws.append_attribute("xmlns") =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    ws.append_child("dimension").append_attribute("ref") = dim_ref.c_str();
    ws.append_child("sheetViews").append_child("sheetView").append_attribute("workbookViewId") = 0;
    ws.append_child("sheetFormatPr").append_attribute("defaultRowHeight") = 15.0;
    auto sheet_data = ws.append_child("sheetData");

    auto inline_str = [](pugi::xml_node row, const std::string& ref, const std::string& text) {
        auto c = row.append_child("c");
        c.append_attribute("r") = ref.c_str();
        c.append_attribute("t") = "inlineStr";
        c.append_child("is").append_child("t").text().set(text.c_str());
    };
    auto number = [](pugi::xml_node row, const std::string& ref, double value) {
        auto c = row.append_child("c");
        c.append_attribute("r") = ref.c_str();
        std::ostringstream ss;
        ss.imbue(std::locale::classic());
        ss << value;
        c.append_child("v").text().set(ss.str().c_str());
    };

    auto header_row = sheet_data.append_child("row");
    header_row.append_attribute("r") = 1;
    for (std::size_t s = 0; s < series.size(); ++s) {
        inline_str(header_row, fmt::format("{}1", column_letter(s + 1)), series[s].name);
    }

    for (std::size_t i = 0; i < categories.size(); ++i) {
        const auto row_num = static_cast<unsigned>(i + 2);
        auto row = sheet_data.append_child("row");
        row.append_attribute("r") = row_num;
        inline_str(row, fmt::format("A{}", row_num), categories[i]);
        for (std::size_t s = 0; s < series.size(); ++s) {
            number(row, fmt::format("{}{}", column_letter(s + 1), row_num), series[s].values[i]);
        }
    }

    std::ostringstream oss;
    doc.save(oss, "", pugi::format_raw);
    return oss.str();
}

} // namespace

void DocxMerger::setChartSeriesName(const std::string& bookmark,
                                    const std::string& old_name,
                                    const std::string& new_name)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setChartSeriesName before load()");
    }

    const std::string chart_part = locate_chart_part(bookmark);
    auto chart_it = parts_.find(chart_part);

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(chart_it->second.data(), chart_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part is not valid XML: {}", chart_part));
    }
    auto plot_area = chart_doc.child("c:chartSpace").child("c:chart").child("c:plotArea");
    if (!plot_area) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no <c:plotArea>");
    }

    bool saw_unsupported = false;
    for (auto type_node : plot_area.children()) {
        const std::string_view tag = type_node.name();
        if (is_unsupported_xy_chart(tag)) {
            saw_unsupported = true;
            continue;
        }
        if (!is_supported_chart_type(tag)) continue;

        for (auto ser : type_node.children("c:ser")) {
            if (read_series_name(ser) != old_name) continue;

            write_series_name(ser, new_name);

            std::ostringstream oss;
            chart_doc.save(oss, "", pugi::format_raw);
            parts_[chart_part] = oss.str();
            return;
        }
    }

    if (saw_unsupported) {
        throw ReportException(
            ReportError::NotImplemented,
            "chart uses scatter/bubble/stock/surface plot — setChartSeriesName only "
            "supports cat/val shapes (bar/line/pie/area/radar/doughnut/3D variants)");
    }
    throw ReportException(
        ReportError::InvalidField,
        fmt::format("series '{}' not found in chart '{}'", old_name, chart_part));
}

void DocxMerger::setChartTitle(const std::string& bookmark, const std::string& title) {
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setChartTitle before load()");
    }

    const std::string chart_part = locate_chart_part(bookmark);
    auto chart_it = parts_.find(chart_part);

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(chart_it->second.data(), chart_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part is not valid XML: {}", chart_part));
    }
    auto chart_node = chart_doc.child("c:chartSpace").child("c:chart");
    auto plot_area  = chart_node.child("c:plotArea");
    if (!plot_area) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no <c:plotArea>");
    }

    bool saw_unsupported = false;
    if (!has_supported_chart_type(plot_area, saw_unsupported) && saw_unsupported) {
        throw ReportException(
            ReportError::NotImplemented,
            "chart uses scatter/bubble/stock/surface plot — setChartTitle only "
            "supports cat/val shapes (bar/line/pie/area/radar/doughnut/3D variants)");
    }

    auto atd_node = chart_node.child("c:autoTitleDeleted");
    set_title(chart_node, atd_node ? atd_node : plot_area, title);
    if (!atd_node) atd_node = chart_node.insert_child_before("c:autoTitleDeleted", plot_area);
    if (auto v = atd_node.attribute("val")) v.set_value("0");
    else atd_node.append_attribute("val") = "0";

    std::ostringstream oss;
    chart_doc.save(oss, "", pugi::format_raw);
    parts_[chart_part] = oss.str();
}

void DocxMerger::setChartAxisTitle(const std::string& bookmark,
                                   ChartAxis           axis,
                                   const std::string&  title)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setChartAxisTitle before load()");
    }

    const std::string chart_part = locate_chart_part(bookmark);
    auto chart_it = parts_.find(chart_part);

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(chart_it->second.data(), chart_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part is not valid XML: {}", chart_part));
    }
    auto plot_area = chart_doc.child("c:chartSpace").child("c:chart").child("c:plotArea");
    if (!plot_area) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no <c:plotArea>");
    }

    bool saw_unsupported = false;
    if (!has_supported_chart_type(plot_area, saw_unsupported) && saw_unsupported) {
        throw ReportException(
            ReportError::NotImplemented,
            "chart uses scatter/bubble/stock/surface plot — setChartAxisTitle only "
            "supports cat/val shapes (bar/line/pie/area/radar/doughnut/3D variants)");
    }

    const char* axis_tag = (axis == ChartAxis::Category) ? "c:catAx" : "c:valAx";
    auto ax = plot_area.child(axis_tag);
    if (!ax) {
        throw ReportException(
            ReportError::InvalidField,
            fmt::format("chart has no {} to set a title on", axis_tag));
    }
    set_title(ax, axis_title_anchor(ax), title);

    std::ostringstream oss;
    chart_doc.save(oss, "", pugi::format_raw);
    parts_[chart_part] = oss.str();
}

void DocxMerger::setChartData(const std::string&              bookmark,
                              const std::vector<std::string>&  categories,
                              const std::vector<ChartSeries>&  series)
{
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "setChartData before load()");
    }
    if (categories.empty()) {
        throw ReportException(ReportError::InvalidField,
                              "setChartData: categories must not be empty");
    }
    if (series.empty()) {
        throw ReportException(ReportError::InvalidField,
                              "setChartData: series must not be empty");
    }
    for (const auto& s : series) {
        if (s.values.size() != categories.size()) {
            throw ReportException(
                ReportError::InvalidField,
                fmt::format("setChartData: series '{}' has {} value(s), expected {} "
                            "(one per category)",
                            s.name, s.values.size(), categories.size()));
        }
    }

    const std::string chart_part = locate_chart_part(bookmark);
    auto chart_it = parts_.find(chart_part);

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(chart_it->second.data(), chart_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("chart part is not valid XML: {}", chart_part));
    }
    auto chart_space = chart_doc.child("c:chartSpace");
    auto plot_area    = chart_space.child("c:chart").child("c:plotArea");
    if (!plot_area) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no <c:plotArea>");
    }

    pugi::xml_node type_node;
    bool saw_unsupported = false;
    for (auto candidate : plot_area.children()) {
        const std::string_view tag = candidate.name();
        if (is_unsupported_xy_chart(tag)) {
            saw_unsupported = true;
            continue;
        }
        if (is_supported_chart_type(tag)) {
            type_node = candidate;
            break;
        }
    }
    if (!type_node) {
        if (saw_unsupported) {
            throw ReportException(
                ReportError::NotImplemented,
                "chart uses scatter/bubble/stock/surface plot — setChartData only "
                "supports cat/val shapes (bar/line/pie/area/radar/doughnut/3D variants)");
        }
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has no recognised series container");
    }

    std::vector<pugi::xml_node> template_sers;
    for (auto ser : type_node.children("c:ser")) template_sers.push_back(ser);
    if (template_sers.empty()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "chart has zero template <c:ser> to clone visual style from");
    }

    // Capture the sheet name from an existing series' range before any
    // rewriting, so the new ranges (and the embedded workbook sync below)
    // stay on the same worksheet the template author picked.
    std::string sheet_name = "Sheet1";
    {
        auto cat = template_sers.front().child("c:cat");
        std::string f = cat.child("c:strRef").child("c:f").child_value();
        if (f.empty()) f = cat.child("c:numRef").child("c:f").child_value();
        const auto parsed = sheet_name_from_formula(f);
        if (!parsed.empty()) sheet_name = parsed;
    }

    // Reconcile the template's series with `series`: reuse existing <c:ser>
    // nodes (and their <c:spPr> styling) in template order, clone the last
    // template series' style for any extra requested series, and drop any
    // surplus template series. Existing <c:dPt> per-point color overrides
    // are dropped below, per series — they're keyed to point indices from
    // the *old* category axis, which have no principled mapping onto a
    // reshaped one (idx 3 might have been "Q4" before and "some other
    // category" now); every point falls back to the chart's normal
    // per-index theme-color rotation instead of carrying over a
    // now-arbitrary override.
    while (template_sers.size() < series.size()) {
        auto clone = type_node.insert_copy_after(template_sers.back(), template_sers.back());
        template_sers.push_back(clone);
    }
    while (template_sers.size() > series.size()) {
        type_node.remove_child(template_sers.back());
        template_sers.pop_back();
    }

    const std::string cat_range =
        fmt::format("{}!$A$2:$A${}", sheet_name, categories.size() + 1);

    for (std::size_t s = 0; s < series.size(); ++s) {
        pugi::xml_node ser = template_sers[s];

        if (auto idx = ser.child("c:idx")) {
            idx.attribute("val").set_value(static_cast<unsigned>(s));
        } else {
            ser.prepend_child("c:idx").append_attribute("val") = static_cast<unsigned>(s);
        }
        if (auto ord = ser.child("c:order")) {
            ord.attribute("val").set_value(static_cast<unsigned>(s));
        } else {
            ser.insert_child_after("c:order", ser.child("c:idx")).append_attribute("val") =
                static_cast<unsigned>(s);
        }

        // Stale per-point color overrides — see the reconciliation comment
        // above for why these can't just be kept.
        while (auto dpt = ser.child("c:dPt")) ser.remove_child(dpt);

        const std::string col       = column_letter(s + 1);
        const std::string name_ref  = fmt::format("{}!${}$1", sheet_name, col);
        const std::string val_range = fmt::format("{}!${}$2:${}${}",
                                                   sheet_name, col, col, categories.size() + 1);

        write_series_name(ser, series[s].name);
        if (auto f = ser.child("c:tx").child("c:strRef").child("c:f")) f.text().set(name_ref.c_str());

        rewrite_string_cache(ser.child("c:cat"), cat_range, categories);
        rewrite_numeric_cache(ser.child("c:val"), val_range, series[s].values);
    }

    std::ostringstream oss;
    chart_doc.save(oss, "", pugi::format_raw);
    parts_[chart_part] = oss.str();

    // ── Sync the embedded workbook, if this chart carries one ──
    const std::string xlsx_part = resolve_chart_external_data(parts_, chart_space, chart_part);
    if (!xlsx_part.empty()) {
        auto xlsx_it = parts_.find(xlsx_part);
        if (xlsx_it == parts_.end()) {
            throw ReportException(
                ReportError::CantCopyDocxTemplate,
                fmt::format("chart references embedded workbook not in archive: {}", xlsx_part));
        }
        auto wb_parts = read_nested_zip(xlsx_it->second);
        const std::string sheet_part = resolve_workbook_sheet_part(wb_parts, sheet_name);
        if (sheet_part.empty()) {
            throw ReportException(
                ReportError::CantCopyDocxTemplate,
                fmt::format("embedded workbook '{}' has no worksheet part", xlsx_part));
        }
        wb_parts[sheet_part] = build_worksheet_xml(categories, series);
        parts_[xlsx_part] = write_nested_zip(wb_parts);
    }
}

// ── paste ───────────────────────────────────────────────────────────────────
//
// v1 semantics: paste() commits the current staged values. Because
// setClipboardValue() writes directly into the DOM, paste() is essentially
// a bookmark-existence check plus a record in pasted_bookmarks() that
// tests can assert against.
//
// A later stage will extend this to clone the section between
// bookmarkStart/bookmarkEnd (needed for table-row and page repetition).

void DocxMerger::paste(const std::string& bookmark) {
    if (!loaded_) {
        throw ReportException(ReportError::CantOpenTemplate,
                              "paste() before load()");
    }
    const auto starts = find_bookmark_starts(bookmark);
    if (starts.empty()) {
        throw ReportException(
            ReportError::InvalidBookmark,
            fmt::format("bookmark not found: {}", bookmark));
    }
    pasted_.push_back(bookmark);
}

} // namespace docweft::docx

// ── factory ────────────────────────────────────────────────────────────────

namespace docweft {

std::unique_ptr<IReportMerger> make_docx_merger() {
    return std::make_unique<docx::DocxMerger>();
}

} // namespace docweft
