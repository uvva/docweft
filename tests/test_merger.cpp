// Tests for docweft::IReportMerger (v1 — Stage 2 of the migration roadmap).
// Covers: load/save, setClipboardValue, paste, and error paths.

#include <catch2/catch_test_macros.hpp>

#include <pugixml.hpp>
#include <zip.h>

#include "docweft/merger.hpp"
#include "docweft/error.hpp"

// Internal header — OK from tests because the test TU is built against
// the same include path as the library and has access to its source tree.
#include "docx/merger.hpp"

#include "fixture_helpers.hpp"

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>
#include <thread>

namespace fs = std::filesystem;

namespace {

// Returns a per-test temp directory that is auto-cleaned by the OS.
fs::path tmp_file(std::string_view stem, std::string_view ext) {
    auto p = fs::temp_directory_path() /
             (std::string("tf_") + std::string(stem) + std::string(ext));
    fs::remove(p);
    return p;
}

// Minimal body with a single bookmark containing a placeholder.
constexpr const char* kBodyWithHeaderBookmark = R"(
    <w:p>
      <w:bookmarkStart w:id="1" w:name="_Header.User"/>
      <w:r><w:rPr/><w:t>Hello {{user}}!</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
)";

// Slurp a single archive entry into a string. Returns empty when absent or
// unreadable — used by tests that need to assert on post-save XML bytes.
inline std::string read_docx_part(const fs::path& docx, const char* name) {
    int err = 0;
    zip_t* z = zip_open(docx.string().c_str(), ZIP_RDONLY, &err);
    if (!z) return {};
    zip_stat_t st{};
    zip_stat_init(&st);
    std::string buf;
    if (zip_stat(z, name, 0, &st) == 0) {
        zip_file_t* f = zip_fopen(z, name, 0);
        if (f) {
            buf.resize(static_cast<std::size_t>(st.size));
            zip_fread(f, buf.data(), st.size);
            zip_fclose(f);
        }
    }
    zip_close(z);
    return buf;
}

// Slurp a single entry out of a zip blob held in memory (e.g. an embedded
// .xlsx read out of read_docx_part's result) via a temp-file round trip —
// same trick DocxMerger uses internally to touch a nested archive.
inline std::string read_nested_part(const std::string& zip_bytes, const char* name) {
    const auto tmp = fs::temp_directory_path() / "tf_nested_read.xlsx";
    {
        std::ofstream f(tmp, std::ios::binary);
        f.write(zip_bytes.data(), static_cast<std::streamsize>(zip_bytes.size()));
    }
    const std::string out = read_docx_part(tmp, name);
    fs::remove(tmp);
    return out;
}

// Body with two bookmarks, including a multi-run field (split placeholder).
constexpr const char* kBodyMultiRun = R"(
    <w:p>
      <w:bookmarkStart w:id="2" w:name="_Header.ReportName"/>
      <w:r><w:t>Report: </w:t></w:r>
      <w:r><w:t>{{title}}</w:t></w:r>
      <w:bookmarkEnd w:id="2"/>
    </w:p>
    <w:p>
      <w:bookmarkStart w:id="3" w:name="_Header.UKDate"/>
      <w:r><w:t>Date {{date}}</w:t></w:r>
      <w:bookmarkEnd w:id="3"/>
    </w:p>
)";

} // namespace

// ── load / save ────────────────────────────────────────────────────────────

TEST_CASE("load rejects missing file", "[merger][load]") {
    auto merger = docweft::make_docx_merger();
    REQUIRE_THROWS_AS(
        merger->load("/path/does/not/exist.docx"),
        docweft::ReportException);

    try { merger->load("/no/such.docx"); }
    catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantOpenTemplate);
    }
}

TEST_CASE("load rejects non-zip file", "[merger][load]") {
    const auto path = tmp_file("not_a_zip", ".docx");
    {
        std::ofstream ofs(path);
        ofs << "this is not a zip archive";
    }

    auto merger = docweft::make_docx_merger();
    REQUIRE_THROWS_AS(merger->load(path.string()), docweft::ReportException);
    fs::remove(path);
}

TEST_CASE("load rejects zip without word/document.xml", "[merger][load]") {
    const auto path = tmp_file("bad_docx", ".docx");
    tf_test::write_docx(path, {
        {"README.txt", "not a real docx"}
    });

    auto merger = docweft::make_docx_merger();
    try {
        merger->load(path.string());
        FAIL("expected ReportException");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantCopyDocxTemplate);
    }
    fs::remove(path);
}

TEST_CASE("load then save produces a valid DOCX round-trip", "[merger][load][save]") {
    const auto in  = tmp_file("roundtrip_in",  ".docx");
    const auto out = tmp_file("roundtrip_out", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    REQUIRE_NOTHROW(merger->load(in.string()));
    REQUIRE_NOTHROW(merger->save(out.string()));

    // The output must be a valid zip that can be re-loaded.
    auto merger2 = docweft::make_docx_merger();
    REQUIRE_NOTHROW(merger2->load(out.string()));

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("save rejects unknown extension", "[merger][save]") {
    const auto in  = tmp_file("ext_in", ".docx");
    const auto out = tmp_file("ext_bad", ".xyz");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->save(out.string());
        FAIL("expected ReportException");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::SaveFailed);
    }
    fs::remove(in);
}

// ── save → .pdf / .html via Word / LibreOffice (Stage 6) ───────────────────

namespace {

// Cross-platform env helper — POSIX setenv / Windows _putenv_s.
void set_env_var(const char* name, const char* value) {
#ifdef _WIN32
    _putenv_s(name, value ? value : "");
#else
    if (value) ::setenv(name, value, 1);
    else       ::unsetenv(name);
#endif
}

// RAII save/restore of a single env var so tests don't leak state.
struct ScopedEnv {
    const char* name;
    bool        had_prev = false;
    std::string prev;
    explicit ScopedEnv(const char* n) : name(n) {
        if (const char* p = std::getenv(name)) { had_prev = true; prev = p; }
    }
    void set(const char* v) { set_env_var(name, v); }
    ~ScopedEnv() {
        if (had_prev) set_env_var(name, prev.c_str());
        else          set_env_var(name, nullptr);
    }
};

bool any_converter_available_on_host() {
#ifdef _WIN32
    if (std::system("reg query HKCR\\Word.Application >nul 2>nul") == 0) return true;
    return std::system("soffice --version >nul 2>nul") == 0;
#else
    return std::system("soffice --version >/dev/null 2>&1") == 0;
#endif
}

} // namespace

TEST_CASE("save('.pdf') returns NoConverter when all converters are disabled",
          "[merger][save][pdf]") {
    const auto in  = tmp_file("pdf_nocv_in",  ".docx");
    const auto out = tmp_file("pdf_nocv_out", ".pdf");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    ScopedEnv kill("DOCWEFT_DISABLE_CONVERTERS");
    kill.set("1");  // forces find_converter() → None regardless of the host

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->save(out.string());
        FAIL("expected NoConverter");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NoConverter);
    }
    fs::remove(in);
}

TEST_CASE("save('.html') returns NoConverter when all converters are disabled",
          "[merger][save][html]") {
    const auto in  = tmp_file("html_nocv_in",  ".docx");
    const auto out = tmp_file("html_nocv_out", ".html");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    ScopedEnv kill("DOCWEFT_DISABLE_CONVERTERS");
    kill.set("1");

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->save(out.string());
        FAIL("expected NoConverter");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NoConverter);
    }
    fs::remove(in);
}

#ifndef _WIN32
// A stand-in for soffice that "converts" by copying its input .docx to
// <outdir>/<stem>.pdf, so the test can inspect exactly what save() handed
// to the converter.
static fs::path write_fake_soffice(std::string_view stem) {
    const auto script = tmp_file(stem, ".sh");
    {
        std::ofstream os(script);
        os << "#!/bin/sh\n"
              "outdir=\n"
              "while [ $# -gt 1 ]; do\n"
              "  [ \"$1\" = --outdir ] && outdir=$2\n"
              "  shift\n"
              "done\n"
              "in=$1\n"
              "cp \"$in\" \"$outdir/$(basename \"$in\" .docx).pdf\"\n";
    }
    fs::permissions(script, fs::perms::owner_all);
    return script;
}

TEST_CASE("save('.pdf') gives the converter a theme when charts have none",
          "[merger][save][pdf][chart]") {
    // Without a theme LibreOffice draws series lacking <c:spPr> with no
    // fill — the chart comes out empty in the PDF.
    const auto in  = tmp_file("pdf_theme_in",  ".docx");
    const auto out = tmp_file("pdf_theme_out", ".pdf");
    const auto out_docx = tmp_file("pdf_theme_out", ".docx");
    tf_test::write_chart_template_docx(in);
    REQUIRE(read_docx_part(in, "word/theme/theme1.xml").empty());

    const auto fake = write_fake_soffice("fake_soffice_theme");
    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    ScopedEnv soffice("DOCWEFT_SOFFICE");
    soffice.set(fake.string().c_str());

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->save(out.string()));

    // `out` is the converter's input, copied verbatim by the fake.
    REQUIRE(read_docx_part(out, "word/theme/theme1.xml").find("accent1") != std::string::npos);
    REQUIRE(read_docx_part(out, "word/_rels/document.xml.rels").find("relationships/theme") != std::string::npos);
    REQUIRE(read_docx_part(out, "[Content_Types].xml").find("/word/theme/theme1.xml") != std::string::npos);

    // The document itself is unchanged: a subsequent .docx save has no theme.
    REQUIRE_NOTHROW(merger->save(out_docx.string()));
    REQUIRE(read_docx_part(out_docx, "word/theme/theme1.xml").empty());

    fs::remove(in);
    fs::remove(out);
    fs::remove(out_docx);
    fs::remove(fake);
}

TEST_CASE("save('.pdf') moves on down the chain when LibreOffice fails",
          "[merger][save][pdf]") {
    const auto in  = tmp_file("pdf_chain_in",  ".docx");
    const auto out = tmp_file("pdf_chain_out", ".pdf");
    const auto fail = tmp_file("fail_soffice", ".sh");
    {
        std::ofstream os(fail);
        os << "#!/bin/sh\nexit 3\n";
    }
    fs::permissions(fail, fs::perms::owner_all);
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    ScopedEnv soffice("DOCWEFT_SOFFICE");
    soffice.set(fail.string().c_str());
    ScopedEnv remote("DOCWEFT_CONVERTER_URL");
    remote.set(nullptr);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->save(out.string());
        // Only possible when the native renderer is compiled in: it is the
        // next link after LibreOffice.
        REQUIRE(fs::file_size(out) > 256);
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::SaveFailed);
        REQUIRE(std::string(e.what()).find("LibreOffice") != std::string::npos);
        REQUIRE_FALSE(fs::exists(out));
    }

    fs::remove(in);
    fs::remove(out);
    fs::remove(fail);
}

TEST_CASE("save('.pdf') stops a hanging LibreOffice after DOCWEFT_SOFFICE_TIMEOUT",
          "[merger][save][pdf]") {
    const auto in   = tmp_file("pdf_hang_in",  ".docx");
    const auto out  = tmp_file("pdf_hang_out", ".html");  // .html: no native fallback
    const auto hang = tmp_file("hang_soffice", ".sh");
    const auto pid_file = tmp_file("hang_soffice", ".pid");
    {
        // Like the real soffice wrapper: the work happens in a child process.
        std::ofstream os(hang);
        os << "#!/bin/sh\n"
              "sleep 60 &\n"
              "echo $! > " << pid_file.string() << "\n"
              "wait\n";
    }
    fs::permissions(hang, fs::perms::owner_all);
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    ScopedEnv soffice("DOCWEFT_SOFFICE");
    soffice.set(hang.string().c_str());
    ScopedEnv timeout("DOCWEFT_SOFFICE_TIMEOUT");
    timeout.set("1");

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    const auto started = std::chrono::steady_clock::now();
    try {
        merger->save(out.string());
        FAIL("expected SaveFailed");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::SaveFailed);
        REQUIRE(std::string(e.what()).find("DOCWEFT_SOFFICE_TIMEOUT") != std::string::npos);
    }
    REQUIRE(std::chrono::steady_clock::now() - started < std::chrono::seconds(10));

    // The child the wrapper started must be gone too, not left running.
    std::ifstream pid_in(pid_file);
    std::string child;
    pid_in >> child;
    REQUIRE_FALSE(child.empty());
    // Poll briefly: a killed orphan stays a zombie until init reaps it.
    const std::string probe = "kill -0 " + child + " 2>/dev/null";
    bool alive = true;
    for (int i = 0; i < 40 && alive; ++i) {
        alive = std::system(probe.c_str()) == 0;
        if (alive) std::this_thread::sleep_for(std::chrono::milliseconds(50));
    }
    REQUIRE_FALSE(alive);

    fs::remove(in);
    fs::remove(hang);
    fs::remove(pid_file);
}

TEST_CASE("save('.html') NoConverter names every converter it checked",
          "[merger][save][html]") {
    const auto in  = tmp_file("html_none_in",  ".docx");
    const auto out = tmp_file("html_none_out", ".html");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    // Neither the native renderer nor the remote converter produce HTML, so
    // with Word and LibreOffice off nothing is left in any build.
    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    ScopedEnv soffice("DOCWEFT_SOFFICE");
    soffice.set("");

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->save(out.string());
        FAIL("expected NoConverter");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NoConverter);
        const std::string msg = e.what();
        REQUIRE(msg.find("Microsoft Word") != std::string::npos);
        REQUIRE(msg.find("LibreOffice") != std::string::npos);
        REQUIRE(msg.find("native PDF renderer") != std::string::npos);
        REQUIRE(msg.find("remote converter") != std::string::npos);
    }
    fs::remove(in);
}

TEST_CASE("save('.pdf') adds no theme to a document without charts",
          "[merger][save][pdf]") {
    const auto in  = tmp_file("pdf_notheme_in",  ".docx");
    const auto out = tmp_file("pdf_notheme_out", ".pdf");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    const auto fake = write_fake_soffice("fake_soffice_notheme");
    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    ScopedEnv soffice("DOCWEFT_SOFFICE");
    soffice.set(fake.string().c_str());

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->save(out.string()));
    REQUIRE(read_docx_part(out, "word/theme/theme1.xml").empty());

    fs::remove(in);
    fs::remove(out);
    fs::remove(fake);
}
#endif

TEST_CASE("save('.pdf') via any available converter produces a non-empty PDF",
          "[merger][save][pdf][integration]") {
    if (!any_converter_available_on_host()) {
        SKIP("no converter available — skipping integration test");
    }
    const auto in  = tmp_file("pdf_ok_in",  ".docx");
    const auto out = tmp_file("pdf_ok_out", ".pdf");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setClipboardValue("_Header.User", "{{user}}", "Alice");
    REQUIRE_NOTHROW(merger->save(out.string()));

    REQUIRE(fs::exists(out));
    REQUIRE(fs::file_size(out) > 256);  // PDF header + trailer minimum

    // %PDF magic at start — both Word and LibreOffice produce valid PDF.
    std::ifstream f(out, std::ios::binary);
    char magic[5] = {};
    f.read(magic, 4);
    REQUIRE(std::string(magic, 4) == "%PDF");

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("save('.pdf') falls back to LibreOffice when MSWord is suppressed",
          "[merger][save][pdf][integration]") {
#ifdef _WIN32
    // On Windows with Word + LibreOffice both present, DOCWEFT_NO_MSWORD
    // must route to LibreOffice. Skip if LibreOffice is missing on host.
    ScopedEnv no_word("DOCWEFT_NO_MSWORD");
    no_word.set("1");
    const bool soffice_ok =
        std::system("soffice --version >nul 2>nul") == 0;
    if (!soffice_ok) {
        SKIP("LibreOffice not available on this Windows host");
    }
#else
    // On non-Windows there is no MSWord branch to suppress — this case
    // reduces to the basic integration test, but we keep the coverage.
    if (std::system("soffice --version >/dev/null 2>&1") != 0) {
        SKIP("soffice not installed");
    }
#endif
    const auto in  = tmp_file("pdf_fb_in",  ".docx");
    const auto out = tmp_file("pdf_fb_out", ".pdf");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->save(out.string()));
    REQUIRE(fs::exists(out));
    REQUIRE(fs::file_size(out) > 256);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("save('.html') via any available converter produces an HTML file",
          "[merger][save][html][integration]") {
    if (!any_converter_available_on_host()) {
        SKIP("no converter available — skipping integration test");
    }
    const auto in  = tmp_file("html_ok_in",  ".docx");
    const auto out = tmp_file("html_ok_out", ".html");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setClipboardValue("_Header.User", "{{user}}", "Борис");
    REQUIRE_NOTHROW(merger->save(out.string()));

    REQUIRE(fs::exists(out));
    REQUIRE(fs::file_size(out) > 64);

    // Both converters write UTF-8; substituted Cyrillic must survive.
    std::ifstream f(out, std::ios::binary);
    const std::string body((std::istreambuf_iterator<char>(f)),
                            std::istreambuf_iterator<char>());
    REQUIRE(body.find("Борис") != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

// ── setCodePage ────────────────────────────────────────────────────────────

TEST_CASE("setCodePage accepts UTF-8", "[merger][codepage]") {
    auto merger = docweft::make_docx_merger();
    REQUIRE_NOTHROW(merger->setCodePage("UTF-8"));
}

TEST_CASE("setCodePage rejects non-UTF-8", "[merger][codepage]") {
    auto merger = docweft::make_docx_merger();
    try {
        merger->setCodePage("CP1251");
        FAIL("expected ReportException");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
}

// ── setClipboardValue ──────────────────────────────────────────────────────

TEST_CASE("setClipboardValue replaces placeholder in bookmark", "[merger][clipboard]") {
    const auto in = tmp_file("clip_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(
        merger->setClipboardValue("_Header.User", "{{user}}", "Alice"));

    // Verify the substitution landed in document.xml.
    // We re-save and re-load to confirm persistence through the archive.
    const auto out = tmp_file("clip_out", ".docx");
    merger->save(out.string());

    auto verifier = docweft::make_docx_merger();
    verifier->load(out.string());

    // Access document.xml via a downcast-free helper: re-save into string
    // by roundtripping through a separate instance... or just check via
    // the public document_xml() on the concrete type. For v1 tests we
    // include the concrete header directly.
    #include "docx/merger.hpp"  // ok for test TU
    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(verifier.get());
    REQUIRE(concrete != nullptr);

    const auto xml = concrete->document_xml();
    REQUIRE(xml.find("Hello Alice!") != std::string::npos);
    REQUIRE(xml.find("{{user}}")     == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setClipboardValue handles placeholder split across runs", "[merger][clipboard]") {
    const auto in = tmp_file("clip_multi_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyMultiRun);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setClipboardValue("_Header.ReportName", "{{title}}", "Q1 Summary");
    merger->setClipboardValue("_Header.UKDate",    "{{date}}",  "2026-04-21");

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    REQUIRE(concrete != nullptr);
    const auto xml = concrete->document_xml();
    REQUIRE(xml.find("Q1 Summary")  != std::string::npos);
    REQUIRE(xml.find("2026-04-21") != std::string::npos);
    REQUIRE(xml.find("{{title}}")  == std::string::npos);
    REQUIRE(xml.find("{{date}}")   == std::string::npos);

    fs::remove(in);
}

TEST_CASE("setClipboardValue preserves UTF-8 / Cyrillic", "[merger][clipboard][utf8]") {
    const auto in = tmp_file("clip_utf8_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setClipboardValue("_Header.User", "{{user}}", "Алексей");

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    const auto xml = concrete->document_xml();
    REQUIRE(xml.find("Алексей") != std::string::npos);

    fs::remove(in);
}

TEST_CASE("setClipboardValue preserves text nodes in overlapping bookmarks",
          "[merger][clipboard][overlap]") {
    // Real-world templates can have overlapping bookmarks: an outer
    // bookmark wraps placeholder text that is *also* inside a nested
    // bookmark for a sibling field. Replacing
    // the outer placeholder must leave the nested one intact, otherwise
    // subsequent setClipboardValue on the nested bookmark hits an empty
    // text span and fails with InvalidField.
    const auto in = tmp_file("overlap_in", ".docx");
    static constexpr const char* kBody = R"(
        <w:p>
          <w:bookmarkStart w:id="10" w:name="outer"/>
          <w:r><w:t>{outer}</w:t></w:r>
          <w:bookmarkStart w:id="11" w:name="inner"/>
          <w:r><w:t>{inner}</w:t></w:r>
          <w:bookmarkEnd w:id="11"/>
          <w:bookmarkEnd w:id="10"/>
        </w:p>
    )";
    tf_test::write_minimal_docx(in, kBody);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setClipboardValue("outer", "{outer}", "Outer!"));
    // Without the targeted-run-only replacement fix, this call used to hit
    // an emptied span and throw InvalidField.
    REQUIRE_NOTHROW(merger->setClipboardValue("inner", "{inner}", "Inner!"));

    const auto out = tmp_file("overlap_out", ".docx");
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto doc = read_docx_part(out, "word/document.xml");
    REQUIRE(doc.find("Outer!") != std::string::npos);
    REQUIRE(doc.find("Inner!") != std::string::npos);
    REQUIRE(doc.find("{outer}") == std::string::npos);
    REQUIRE(doc.find("{inner}") == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setClipboardValue fills _Header / _Footer parts that carry no bookmarks",
          "[merger][clipboard][header]") {
    // word/header{N}.xml and word/footer{N}.xml are often never bookmarked —
    // document.xml only reaches them via section properties. "_Header"/"_Footer" are
    // pseudo-bookmarks that address "every header/footer part" instead of a
    // real <w:bookmarkStart>.
    const auto in = tmp_file("header_footer_in", ".docx");

    static constexpr const char* kContentTypes = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>)";
    static constexpr const char* kPackageRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>)";
    static constexpr const char* kDocRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>)";
    static constexpr const char* kDocument = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body>
</w:document>)";
    // Mirrors the real report.docx: {User} in one run, {ReportName} split
    // across three runs with an intervening <w:proofErr> — no bookmark.
    static constexpr const char* kHeader1 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>{User}</w:t></w:r>
    <w:r><w:t>{</w:t></w:r>
    <w:proofErr w:type="spellStart"/>
    <w:r><w:t>ReportName</w:t></w:r>
    <w:proofErr w:type="spellEnd"/>
    <w:r><w:t>}</w:t></w:r>
  </w:p>
</w:hdr>)";
    // A second, unrelated header part (e.g. "first page" header) that must
    // stay completely untouched.
    static constexpr const char* kHeader2 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>First-page header, no placeholders here</w:t></w:r></w:p>
</w:hdr>)";
    static constexpr const char* kFooter1 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>{kPAGE}</w:t></w:r></w:p>
</w:ftr>)";

    tf_test::write_docx(in, {
        {"[Content_Types].xml",          kContentTypes},
        {"_rels/.rels",                  kPackageRels},
        {"word/_rels/document.xml.rels", kDocRels},
        {"word/document.xml",            kDocument},
        {"word/header1.xml",             kHeader1},
        {"word/header2.xml",             kHeader2},
        {"word/footer1.xml",             kFooter1},
    });

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setClipboardValue("_Header", "{User}", "Operator"));
    REQUIRE_NOTHROW(merger->setClipboardValue("_Header", "{ReportName}",
                                              "ALTAMI STUDIO REPORT"));
    REQUIRE_NOTHROW(merger->setClipboardValue("_Footer", "{kPAGE}", "Page 1 of 3"));

    const auto out = tmp_file("header_footer_out", ".docx");
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto h1 = read_docx_part(out, "word/header1.xml");
    REQUIRE(h1.find("Operator") != std::string::npos);
    REQUIRE(h1.find("ALTAMI STUDIO REPORT") != std::string::npos);
    REQUIRE(h1.find("{User}") == std::string::npos);
    REQUIRE(h1.find("{ReportName}") == std::string::npos);

    // header2.xml matched neither field — must be byte-for-byte untouched.
    REQUIRE(read_docx_part(out, "word/header2.xml") == kHeader2);

    const auto f1 = read_docx_part(out, "word/footer1.xml");
    REQUIRE(f1.find("Page 1 of 3") != std::string::npos);
    REQUIRE(f1.find("{kPAGE}") == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setClipboardValue on _Header throws InvalidBookmark when the archive has no header parts",
          "[merger][clipboard][header][error]") {
    const auto in = tmp_file("header_none_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setClipboardValue("_Header", "{User}", "x");
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setClipboardValue on _Header throws InvalidField when no header part has the placeholder",
          "[merger][clipboard][header][error]") {
    const auto in = tmp_file("header_field_err_in", ".docx");
    static constexpr const char* kContentTypes = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>)";
    static constexpr const char* kPackageRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>)";
    static constexpr const char* kDocRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>)";
    static constexpr const char* kDocument = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body>
</w:document>)";
    static constexpr const char* kHeader1 = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Nothing to see here</w:t></w:r></w:p>
</w:hdr>)";

    tf_test::write_docx(in, {
        {"[Content_Types].xml",          kContentTypes},
        {"_rels/.rels",                  kPackageRels},
        {"word/_rels/document.xml.rels", kDocRels},
        {"word/document.xml",            kDocument},
        {"word/header1.xml",             kHeader1},
    });

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setClipboardValue("_Header", "{User}", "x");
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setClipboardValue throws on unknown bookmark", "[merger][clipboard][error]") {
    const auto in = tmp_file("clip_err_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setClipboardValue("_NoSuch.Bookmark", "{{user}}", "x");
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setClipboardValue throws on unknown field", "[merger][clipboard][error]") {
    const auto in = tmp_file("clip_field_err_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setClipboardValue("_Header.User", "{{nope}}", "x");
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

// ── setTableRow (Stage 3) ──────────────────────────────────────────────────

namespace {

// Count occurrences of `needle` in `haystack` (ripgrep-free).
std::size_t count_substr(std::string_view haystack, std::string_view needle) {
    if (needle.empty()) return 0;
    std::size_t n = 0;
    std::size_t pos = 0;
    while ((pos = haystack.find(needle, pos)) != std::string_view::npos) {
        ++n;
        pos += needle.size();
    }
    return n;
}

} // namespace

TEST_CASE("setTableRow clones template row per data row", "[merger][table]") {
    const auto in = tmp_file("tbl_three", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    merger->setTableRow("TableRow",
        {"{{name}}", "{{score}}"},
        {
            {"Alice", "95"},
            {"Bob",   "82"},
            {"Carol", "77"},
        });

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    REQUIRE(concrete != nullptr);
    const auto xml = concrete->document_xml();

    // Header row + 3 cloned rows. Template row is removed.
    REQUIRE(count_substr(xml, "<w:tr>") == 4);
    REQUIRE(xml.find("Alice") != std::string::npos);
    REQUIRE(xml.find("Bob")   != std::string::npos);
    REQUIRE(xml.find("Carol") != std::string::npos);
    REQUIRE(xml.find("{{name}}")  == std::string::npos);
    REQUIRE(xml.find("{{score}}") == std::string::npos);
    // Header text survives the row cloning.
    REQUIRE(xml.find("Name")  != std::string::npos);
    REQUIRE(xml.find("Score") != std::string::npos);

    fs::remove(in);
}

TEST_CASE("setTableRow with empty rows drops the template row", "[merger][table]") {
    const auto in = tmp_file("tbl_empty", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setTableRow("TableRow", {"{{name}}", "{{score}}"}, {});

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    const auto xml = concrete->document_xml();
    // Only the header row remains.
    REQUIRE(count_substr(xml, "<w:tr>") == 1);
    REQUIRE(xml.find("{{name}}")  == std::string::npos);
    REQUIRE(xml.find("{{score}}") == std::string::npos);

    fs::remove(in);
}

TEST_CASE("setTableRow strips bookmarks from clones (no duplicate w:id)",
          "[merger][table]") {
    const auto in = tmp_file("tbl_bmk", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setTableRow("TableRow",
        {"{{name}}", "{{score}}"},
        {{"Alice", "95"}, {"Bob", "82"}});

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    const auto xml = concrete->document_xml();
    // Cloned rows must not carry the original bookmark — otherwise Word
    // refuses to open the file on duplicate w:id.
    REQUIRE(xml.find("w:bookmarkStart") == std::string::npos);
    REQUIRE(xml.find("w:bookmarkEnd")   == std::string::npos);

    fs::remove(in);
}

TEST_CASE("setTableRow round-trips through save/load", "[merger][table][save]") {
    const auto in  = tmp_file("tbl_rt_in",  ".docx");
    const auto out = tmp_file("tbl_rt_out", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    merger->setTableRow("TableRow",
        {"{{name}}", "{{score}}"},
        {{"Алиса", "95"}, {"Борис", "82"}});
    merger->save(out.string());

    auto verifier = docweft::make_docx_merger();
    verifier->load(out.string());
    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(verifier.get());
    const auto xml = concrete->document_xml();
    REQUIRE(xml.find("Алиса") != std::string::npos);
    REQUIRE(xml.find("Борис") != std::string::npos);
    REQUIRE(count_substr(xml, "<w:tr>") == 3);  // header + 2 data rows

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setTableRow throws on unknown bookmark", "[merger][table][error]") {
    const auto in = tmp_file("tbl_nobmk", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setTableRow("NoSuch", {"{{name}}"}, {{"x"}});
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setTableRow throws when bookmark is not in a <w:tr>",
          "[merger][table][error]") {
    // Reuse the plain header-bookmark fixture — bookmark sits inside a <w:p>
    // that has no enclosing <w:tr>.
    const auto in = tmp_file("tbl_not_tr", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setTableRow("_Header.User", {"{{user}}"}, {{"Alice"}});
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setTableRow throws on row width mismatch", "[merger][table][error]") {
    const auto in = tmp_file("tbl_mismatch", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        merger->setTableRow("TableRow",
            {"{{name}}", "{{score}}"},
            {{"Alice", "95"}, {"Bob"}});  // second row is too short
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setTableRow throws when a field is absent from the row template",
          "[merger][table][error]") {
    const auto in = tmp_file("tbl_bad_field", ".docx");
    tf_test::write_table_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());

    try {
        // {{nope}} is not present in any cell of the template row.
        merger->setTableRow("TableRow", {"{{nope}}"}, {{"x"}});
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

// ── setImage (Stage 5) ─────────────────────────────────────────────────────

namespace {

constexpr const char* kBodyWithImageBookmark = R"(
    <w:p>
      <w:bookmarkStart w:id="5" w:name="Body.Image"/>
      <w:r><w:t>{{image}}</w:t></w:r>
      <w:bookmarkEnd w:id="5"/>
    </w:p>
)";

} // namespace

TEST_CASE("setImage inserts a <w:drawing> and registers the relationship",
          "[merger][image]") {
    const auto docx = tmp_file("img_basic_in", ".docx");
    const auto png  = tmp_file("img_basic",    ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    REQUIRE_NOTHROW(merger->setImage("Body.Image", png));

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    REQUIRE(concrete != nullptr);
    const auto xml = concrete->document_xml();

    // Drawing element landed with the rId and the 2×2 PNG's EMU extents.
    // 2 px × 9525 EMU/px = 19050.
    REQUIRE(xml.find("<w:drawing>") != std::string::npos);
    REQUIRE(xml.find("cx=\"19050\"") != std::string::npos);
    REQUIRE(xml.find("cy=\"19050\"") != std::string::npos);
    // Text placeholder has been displaced by the picture.
    REQUIRE(xml.find("{{image}}") == std::string::npos);

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage saves a round-trippable DOCX with embedded PNG",
          "[merger][image][save]") {
    const auto in_docx  = tmp_file("img_rt_in",  ".docx");
    const auto out_docx = tmp_file("img_rt_out", ".docx");
    const auto png      = tmp_file("img_rt",     ".png");
    tf_test::write_minimal_docx(in_docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(in_docx.string());
    merger->setImage("Body.Image", png);
    REQUIRE_NOTHROW(merger->save(out_docx.string()));

    // Re-load the saved .docx and inspect the archive via a fresh merger.
    auto verifier = docweft::make_docx_merger();
    REQUIRE_NOTHROW(verifier->load(out_docx.string()));

    // Pull the raw archive via libzip to validate media + rels + content-types.
    int err = 0;
    zip_t* z = zip_open(out_docx.string().c_str(), ZIP_RDONLY, &err);
    REQUIRE(z != nullptr);

    auto read_part = [&](const char* name) -> std::string {
        zip_stat_t st{};
        zip_stat_init(&st);
        if (zip_stat(z, name, 0, &st) != 0) return {};
        zip_file_t* f = zip_fopen(z, name, 0);
        if (!f) return {};
        std::string buf;
        buf.resize(static_cast<std::size_t>(st.size));
        zip_fread(f, buf.data(), st.size);
        zip_fclose(f);
        return buf;
    };

    const auto media = read_part("word/media/image1.png");
    REQUIRE(!media.empty());
    REQUIRE(media.size() == tf_test::tiny_png_bytes().size());
    REQUIRE(std::memcmp(media.data(), tf_test::tiny_png_bytes().data(),
                        media.size()) == 0);

    const auto rels = read_part("word/_rels/document.xml.rels");
    REQUIRE(rels.find("Target=\"media/image1.png\"") != std::string::npos);
    REQUIRE(rels.find("/relationships/image") != std::string::npos);

    const auto ct = read_part("[Content_Types].xml");
    REQUIRE(ct.find("Extension=\"png\"")  != std::string::npos);
    REQUIRE(ct.find("image/png")          != std::string::npos);

    zip_close(z);

    fs::remove(in_docx);
    fs::remove(out_docx);
    fs::remove(png);
}

TEST_CASE("setImage preserves existing content-type entries",
          "[merger][image]") {
    const auto docx = tmp_file("img_ct_in", ".docx");
    const auto png  = tmp_file("img_ct",    ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    merger->setImage("Body.Image", png);
    // Call again with the same PNG to ensure idempotency of content-type insert.
    merger->setImage("Body.Image", png);

    const auto out = tmp_file("img_ct_out", ".docx");
    merger->save(out.string());

    int err = 0;
    zip_t* z = zip_open(out.string().c_str(), ZIP_RDONLY, &err);
    REQUIRE(z != nullptr);
    zip_stat_t st{}; zip_stat_init(&st);
    REQUIRE(zip_stat(z, "[Content_Types].xml", 0, &st) == 0);
    std::string ct_xml(static_cast<std::size_t>(st.size), '\0');
    zip_file_t* f = zip_fopen(z, "[Content_Types].xml", 0);
    zip_fread(f, ct_xml.data(), st.size);
    zip_fclose(f);
    zip_close(z);

    // exactly one <Default Extension="png" ...> — no duplication on second call
    std::size_t count = 0;
    std::size_t pos = 0;
    while ((pos = ct_xml.find("Extension=\"png\"", pos)) != std::string::npos) {
        ++count;
        pos += 1;
    }
    REQUIRE(count == 1);

    fs::remove(docx);
    fs::remove(png);
    fs::remove(out);
}

TEST_CASE("setImage allocates unique media paths for multiple calls",
          "[merger][image]") {
    const auto docx = tmp_file("img_multi_in", ".docx");
    const auto png  = tmp_file("img_multi",    ".png");
    // Two separate bookmarks so each call targets a different paragraph.
    constexpr const char* kBody = R"(
        <w:p>
          <w:bookmarkStart w:id="5" w:name="Body.One"/>
          <w:r><w:t>{{one}}</w:t></w:r>
          <w:bookmarkEnd w:id="5"/>
        </w:p>
        <w:p>
          <w:bookmarkStart w:id="6" w:name="Body.Two"/>
          <w:r><w:t>{{two}}</w:t></w:r>
          <w:bookmarkEnd w:id="6"/>
        </w:p>
    )";
    tf_test::write_minimal_docx(docx, kBody);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    merger->setImage("Body.One", png);
    merger->setImage("Body.Two", png);

    const auto out = tmp_file("img_multi_out", ".docx");
    merger->save(out.string());

    int err = 0;
    zip_t* z = zip_open(out.string().c_str(), ZIP_RDONLY, &err);
    REQUIRE(z != nullptr);
    zip_stat_t s1{}, s2{};
    zip_stat_init(&s1);
    zip_stat_init(&s2);
    REQUIRE(zip_stat(z, "word/media/image1.png", 0, &s1) == 0);
    REQUIRE(zip_stat(z, "word/media/image2.png", 0, &s2) == 0);
    zip_close(z);

    fs::remove(docx);
    fs::remove(png);
    fs::remove(out);
}

TEST_CASE("setImage throws InvalidBookmark on unknown bookmark",
          "[merger][image][error]") {
    const auto docx = tmp_file("img_err_bmk", ".docx");
    const auto png  = tmp_file("img_err_bmk", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.NotHere", png);
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage throws CantOpenTemplate when the image file is missing",
          "[merger][image][error]") {
    const auto docx = tmp_file("img_err_file", ".docx");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.Image", "/no/such/image.png");
        FAIL("expected CantOpenTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantOpenTemplate);
    }
    fs::remove(docx);
}

// ── setImage with ImageSize bounds ─────────────────────────────────────────

namespace {

// Extract the first `cx="..."` and `cy="..."` from a <wp:extent>, returning
// the numeric EMU values.
void extract_extent(const std::string& xml,
                    std::uint64_t& cx, std::uint64_t& cy) {
    const auto extent = xml.find("<wp:extent ");
    REQUIRE(extent != std::string::npos);
    const auto cx_pos = xml.find("cx=\"", extent);
    const auto cy_pos = xml.find("cy=\"", extent);
    REQUIRE(cx_pos != std::string::npos);
    REQUIRE(cy_pos != std::string::npos);
    cx = std::stoull(xml.substr(cx_pos + 4,
                                xml.find('"', cx_pos + 4) - cx_pos - 4));
    cy = std::stoull(xml.substr(cy_pos + 4,
                                xml.find('"', cy_pos + 4) - cy_pos - 4));
}

constexpr std::uint32_t kEmuPerPx = 9525;

} // namespace

TEST_CASE("setImage with default ImageSize{} uses intrinsic pixel size",
          "[merger][image][size]") {
    const auto docx = tmp_file("img_sz_def", ".docx");
    const auto png  = tmp_file("img_sz_def", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);  // 2×2

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    merger->setImage("Body.Image", png);  // no bounds

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    std::uint64_t cx = 0, cy = 0;
    extract_extent(concrete->document_xml(), cx, cy);
    REQUIRE(cx == 2 * kEmuPerPx);
    REQUIRE(cy == 2 * kEmuPerPx);

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage downscales when max bound is smaller than native",
          "[merger][image][size]") {
    const auto docx = tmp_file("img_sz_max", ".docx");
    const auto png  = tmp_file("img_sz_max", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);  // 2×2

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    // max_width_px=1 → scale must become 0.5 → displayed extent 1×1 px.
    merger->setImage("Body.Image", png, docweft::ImageSize{0, 0, 1, 0});

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    std::uint64_t cx = 0, cy = 0;
    extract_extent(concrete->document_xml(), cx, cy);
    REQUIRE(cx == kEmuPerPx);  // 1 px worth of EMU
    REQUIRE(cy == kEmuPerPx);  // aspect preserved

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage upscales when min bound is larger than native",
          "[merger][image][size]") {
    const auto docx = tmp_file("img_sz_min", ".docx");
    const auto png  = tmp_file("img_sz_min", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);  // 2×2

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    // min_width_px=10 → scale=5, displayed extent 10×10 px.
    merger->setImage("Body.Image", png, docweft::ImageSize{10, 0, 0, 0});

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    std::uint64_t cx = 0, cy = 0;
    extract_extent(concrete->document_xml(), cx, cy);
    REQUIRE(cx == 10 * kEmuPerPx);
    REQUIRE(cy == 10 * kEmuPerPx);

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage keeps intrinsic size when native fits inside bounds",
          "[merger][image][size]") {
    const auto docx = tmp_file("img_sz_fit", ".docx");
    const auto png  = tmp_file("img_sz_fit", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);  // 2×2

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    // 1×1 ≤ 2×2 ≤ 100×100 — intrinsic scale=1 is in the band, no change.
    merger->setImage("Body.Image", png,
                     docweft::ImageSize{1, 1, 100, 100});

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    std::uint64_t cx = 0, cy = 0;
    extract_extent(concrete->document_xml(), cx, cy);
    REQUIRE(cx == 2 * kEmuPerPx);
    REQUIRE(cy == 2 * kEmuPerPx);

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage throws InvalidField on contradictory bounds",
          "[merger][image][size][error]") {
    const auto docx = tmp_file("img_sz_bad", ".docx");
    const auto png  = tmp_file("img_sz_bad", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);  // 2×2

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        // min_width 20 → scale ≥ 10; max_height 5 → scale ≤ 2.5. Impossible.
        merger->setImage("Body.Image", png,
                         docweft::ImageSize{20, 0, 0, 5});
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }

    fs::remove(docx);
    fs::remove(png);
}

TEST_CASE("setImage leaves document unmutated on contradictory bounds",
          "[merger][image][size][error]") {
    // Confirms the compute_display_scale check happens BEFORE any archive
    // mutation — media part, relationship, content-type should not leak.
    const auto docx = tmp_file("img_sz_atomic", ".docx");
    const auto png  = tmp_file("img_sz_atomic", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_png(png);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.Image", png,
                         docweft::ImageSize{20, 0, 0, 5});
    } catch (const docweft::ReportException&) {
        // Expected — fall through to verification.
    }

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    const auto xml = concrete->document_xml();
    // Original {{image}} placeholder must still be there; no <w:drawing>
    // inserted; no new relationship.
    REQUIRE(xml.find("{{image}}") != std::string::npos);
    REQUIRE(xml.find("<w:drawing>") == std::string::npos);

    fs::remove(docx);
    fs::remove(png);
}

#if defined(DOCWEFT_HAVE_STB)
TEST_CASE("setImage transcodes BMP to PNG (Phase 2, stb path)",
          "[merger][image][bmp]") {
    const auto docx = tmp_file("img_bmp_in",  ".docx");
    const auto bmp  = tmp_file("img_bmp_src", ".bmp");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_bmp(bmp);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());
    REQUIRE_NOTHROW(merger->setImage("Body.Image", bmp));

    // Archive should contain a PNG, not the original BMP — transcoded inline.
    const auto out = tmp_file("img_bmp_out", ".docx");
    merger->save(out.string());

    int err = 0;
    zip_t* z = zip_open(out.string().c_str(), ZIP_RDONLY, &err);
    REQUIRE(z != nullptr);
    zip_stat_t st{}; zip_stat_init(&st);
    REQUIRE(zip_stat(z, "word/media/image1.png", 0, &st) == 0);
    REQUIRE(st.size > 0);
    zip_file_t* f = zip_fopen(z, "word/media/image1.png", 0);
    REQUIRE(f != nullptr);
    char magic[8] = {};
    zip_fread(f, magic, 8);
    zip_fclose(f);
    zip_close(z);

    // PNG signature on the stored bytes — confirms the BMP was re-encoded.
    static constexpr unsigned char kPngSig[8] = {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A};
    REQUIRE(std::memcmp(magic, kPngSig, 8) == 0);

    fs::remove(docx);
    fs::remove(bmp);
    fs::remove(out);
}
#else
TEST_CASE("setImage rejects non-PNG when stb decoder is not compiled in",
          "[merger][image][error]") {
    const auto docx = tmp_file("img_no_stb", ".docx");
    const auto bmp  = tmp_file("img_no_stb", ".bmp");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    tf_test::write_tiny_bmp(bmp);

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.Image", bmp);
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(docx);
    fs::remove(bmp);
}
#endif

#if defined(DOCWEFT_HAVE_STB)
TEST_CASE("setImage rejects corrupt JPEG bytes with CantCopyDocxTemplate",
          "[merger][image][error]") {
    const auto docx = tmp_file("img_bad_jpg", ".docx");
    const auto bad  = tmp_file("img_bad_jpg", ".jpg");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    {
        // Valid JPEG magic FFD8 followed by garbage — stb_image refuses it.
        std::ofstream f(bad, std::ios::binary);
        const unsigned char head[] = {0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 'j', 'U', 'N', 'K'};
        f.write(reinterpret_cast<const char*>(head), sizeof(head));
    }

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.Image", bad);
        FAIL("expected CantCopyDocxTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantCopyDocxTemplate);
    }
    fs::remove(docx);
    fs::remove(bad);
}
#endif

TEST_CASE("setImage rejects non-image files",
          "[merger][image][error]") {
    const auto docx = tmp_file("img_err_junk", ".docx");
    const auto junk = tmp_file("img_err_junk", ".png");
    tf_test::write_minimal_docx(docx, kBodyWithImageBookmark);
    {
        std::ofstream f(junk, std::ios::binary);
        f << "this is not any known image format";
    }

    auto merger = docweft::make_docx_merger();
    merger->load(docx.string());

    try {
        merger->setImage("Body.Image", junk);
        FAIL("expected CantCopyDocxTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantCopyDocxTemplate);
    }
    fs::remove(docx);
    fs::remove(junk);
}

// ── setChartValue (Stage 4) ────────────────────────────────────────────────

TEST_CASE("setChartValue rewrites a bar chart numeric cache",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_bar_in",  ".docx");
    const auto out = tmp_file("chart_bar_out", ".docx");
    tf_test::write_chart_template_docx(in);   // Bar, 2×2, Sales/Costs × Q1/Q2

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartValue(
        "kStatChart", "ChartField", "Sales", "Q1", 99.0));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto chart = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(!chart.empty());
    // The Sales/Q1 slot was 10 initially; must now be 99.
    REQUIRE(chart.find(R"(<c:pt idx="0"><c:v>99</c:v></c:pt>)")
            != std::string::npos);
    // Sales/Q2 must remain 20 (untouched).
    REQUIRE(chart.find(R"(<c:pt idx="1"><c:v>20</c:v></c:pt>)")
            != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartValue updates the second series in a multi-series chart",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_multi_in",  ".docx");
    const auto out = tmp_file("chart_multi_out", ".docx");
    tf_test::write_chart_template_docx(in);   // Costs row: {5, 15}

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartValue(
        "kStatChart", "ChartField", "Costs", "Q2", 77.5));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto chart = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(!chart.empty());
    // 77.5 in locale-independent form
    REQUIRE(chart.find(R"(<c:pt idx="1"><c:v>77.5</c:v></c:pt>)")
            != std::string::npos);
    // First series untouched: Sales/Q2 remains 20.
    REQUIRE(chart.find(R"(<c:v>20</c:v>)") != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartValue survives a round-trip through save/load",
          "[merger][chart]") {
    const auto in   = tmp_file("chart_rt_in",  ".docx");
    const auto out1 = tmp_file("chart_rt_m1",  ".docx");
    const auto out2 = tmp_file("chart_rt_m2",  ".docx");
    tf_test::write_chart_template_docx(in);

    auto m1 = docweft::make_docx_merger();
    m1->load(in.string());
    m1->setChartValue("kStatChart", "ChartField", "Sales", "Q1", 42.0);
    m1->save(out1.string());

    // Re-open and mutate again — the value we wrote must be what we see.
    auto m2 = docweft::make_docx_merger();
    m2->load(out1.string());
    m2->setChartValue("kStatChart", "ChartField", "Sales", "Q2", 43.0);
    m2->save(out2.string());

    const auto chart = read_docx_part(out2, "word/charts/chart1.xml");
    REQUIRE(chart.find(R"(<c:pt idx="0"><c:v>42</c:v></c:pt>)")
            != std::string::npos);
    REQUIRE(chart.find(R"(<c:pt idx="1"><c:v>43</c:v></c:pt>)")
            != std::string::npos);

    fs::remove(in);
    fs::remove(out1);
    fs::remove(out2);
}

TEST_CASE("setChartValue preserves Cyrillic series and category names",
          "[merger][chart][utf8]") {
    const auto in  = tmp_file("chart_utf8_in",  ".docx");
    const auto out = tmp_file("chart_utf8_out", ".docx");

    tf_test::ChartFixture fx;
    fx.series_names   = {"Продажи", "Расходы"};
    fx.category_names = {"Январь",  "Февраль"};
    fx.data           = {{100, 200}, {50, 60}};
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartValue(
        "kStatChart", "ChartField", "Продажи", "Февраль", 250.0));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto chart = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(!chart.empty());
    REQUIRE(chart.find("Продажи") != std::string::npos);
    REQUIRE(chart.find("Февраль") != std::string::npos);
    REQUIRE(chart.find(R"(<c:pt idx="1"><c:v>250</c:v></c:pt>)")
            != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartValue works on a line chart", "[merger][chart]") {
    const auto in  = tmp_file("chart_line_in",  ".docx");
    const auto out = tmp_file("chart_line_out", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Line;
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartValue(
        "kStatChart", "ChartField", "Sales", "Q1", 33.0));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto chart = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(chart.find("c:lineChart") != std::string::npos);
    REQUIRE(chart.find(R"(<c:pt idx="0"><c:v>33</c:v></c:pt>)")
            != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartValue works on a pie chart", "[merger][chart]") {
    const auto in  = tmp_file("chart_pie_in",  ".docx");
    const auto out = tmp_file("chart_pie_out", ".docx");
    tf_test::ChartFixture fx;
    fx.kind          = tf_test::ChartKind::Pie;
    fx.series_names  = {"Share"};      // pie has one series
    fx.data          = {{30, 70}};
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartValue(
        "kStatChart", "ChartField", "Share", "Q2", 65.0));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto chart = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(chart.find("c:pieChart") != std::string::npos);
    REQUIRE(chart.find(R"(<c:pt idx="1"><c:v>65</c:v></c:pt>)")
            != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartValue throws InvalidBookmark on unknown bookmark",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_nobmk_in", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartValue("NoSuchBmk", "f", "Sales", "Q1", 1.0);
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setChartValue throws InvalidBookmark when bookmark has no chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_nopara_in", ".docx");
    // Bookmark in a plain text paragraph, no drawing anywhere.
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartValue("_Header.User", "f", "Sales", "Q1", 1.0);
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

TEST_CASE("setChartValue throws InvalidField on unknown series",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_noser_in", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartValue("kStatChart", "f", "NoSuchSeries", "Q1", 1.0);
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartValue throws InvalidField on unknown category",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_nocat_in", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartValue("kStatChart", "f", "Sales", "NoSuchCat", 1.0);
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartValue throws NotImplemented on scatter chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_scatter_in", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Scatter;
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartValue("kStatChart", "f", "Sales", "Q1", 1.0);
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(in);
}

TEST_CASE("setChartValue before load throws CantOpenTemplate",
          "[merger][chart][error]") {
    auto merger = docweft::make_docx_merger();
    try {
        merger->setChartValue("b", "f", "s", "c", 1.0);
        FAIL("expected CantOpenTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantOpenTemplate);
    }
}

// ── setChartSeriesName ────────────────────────────────────────────────────

TEST_CASE("setChartSeriesName renames a series without touching its data",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_rename_in",  ".docx");
    const auto out = tmp_file("chart_rename_out", ".docx");
    tf_test::write_chart_template_docx(in);   // Sales/Costs × Q1/Q2

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartSeriesName("kStatChart", "Sales", "Revenue"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto bar = doc.child("c:chartSpace").child("c:chart").child("c:plotArea").child("c:barChart");
    auto sers = bar.children("c:ser");
    auto it = sers.begin();
    REQUIRE(std::string((*it).child("c:tx").child("c:strRef").child("c:strCache")
                         .child("c:pt").child_value("c:v")) == "Revenue");
    ++it;
    REQUIRE(std::string((*it).child("c:tx").child("c:strRef").child("c:strCache")
                         .child("c:pt").child_value("c:v")) == "Costs");
    // Values untouched.
    REQUIRE(xml.find(R"(<c:v>10</c:v>)") != std::string::npos);
    REQUIRE(xml.find(R"(<c:v>20</c:v>)") != std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartSeriesName throws InvalidField on unknown series",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_rename_bad", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartSeriesName("kStatChart", "NoSuchSeries", "X");
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartSeriesName throws NotImplemented on scatter chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_rename_scatter", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Scatter;
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartSeriesName("kStatChart", "Sales", "X");
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(in);
}

TEST_CASE("setChartSeriesName before load throws CantOpenTemplate",
          "[merger][chart][error]") {
    auto merger = docweft::make_docx_merger();
    try {
        merger->setChartSeriesName("b", "s", "n");
        FAIL("expected CantOpenTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantOpenTemplate);
    }
}

// ── setChartTitle ─────────────────────────────────────────────────────────

TEST_CASE("setChartTitle creates a title and clears autoTitleDeleted",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_title_in",  ".docx");
    const auto out = tmp_file("chart_title_out", ".docx");
    tf_test::write_chart_template_docx(in);   // no <c:title> in the fixture

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartTitle("kStatChart", "Quarterly Sales"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto chart_node = doc.child("c:chartSpace").child("c:chart");
    auto title_text = chart_node.child("c:title").child("c:tx").child("c:rich")
                       .child("a:p").child("a:r").child_value("a:t");
    REQUIRE(std::string(title_text) == "Quarterly Sales");
    REQUIRE(std::string(chart_node.child("c:autoTitleDeleted").attribute("val").value()) == "0");

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartTitle replaces an existing title instead of duplicating it",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_title_replace_in",  ".docx");
    const auto out = tmp_file("chart_title_replace_out", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartTitle("kStatChart", "First Title"));
    REQUIRE_NOTHROW(merger->setChartTitle("kStatChart", "Second Title"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(xml.find("First Title") == std::string::npos);
    REQUIRE(xml.find("Second Title") != std::string::npos);
    // Exactly one <a:t> run under the title — no leftover duplicate runs.
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto title = doc.child("c:chartSpace").child("c:chart").child("c:title");
    std::size_t run_count = 0;
    for (auto r : title.child("c:tx").child("c:rich").child("a:p").children("a:r")) {
        (void)r;
        ++run_count;
    }
    REQUIRE(run_count == 1);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartTitle preserves the template's existing title formatting",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_title_fmt_in",  ".docx");
    const auto out = tmp_file("chart_title_fmt_out", ".docx");
    tf_test::ChartFixture fx;
    fx.initial_title = "Old Title";   // fixture styles this red/bold/14pt/centered/rotated
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartTitle("kStatChart", "New Title"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto title = doc.child("c:chartSpace").child("c:chart").child("c:title");
    auto rich  = title.child("c:tx").child("c:rich");
    auto run   = rich.child("a:p").child("a:r");

    REQUIRE(std::string(run.child_value("a:t")) == "New Title");
    // The template's box/paragraph/run formatting survived the text swap.
    REQUIRE(std::string(rich.child("a:bodyPr").attribute("rot").value()) == "5400000");
    REQUIRE(std::string(rich.child("a:p").child("a:pPr").attribute("algn").value()) == "ctr");
    REQUIRE(std::string(run.child("a:rPr").attribute("b").value()) == "1");
    REQUIRE(std::string(run.child("a:rPr").attribute("sz").value()) == "1400");
    REQUIRE(std::string(run.child("a:rPr").child("a:solidFill").child("a:srgbClr")
                         .attribute("val").value()) == "FF0000");

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartTitle throws NotImplemented on scatter chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_title_scatter", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Scatter;
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartTitle("kStatChart", "X");
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(in);
}

// ── setChartAxisTitle ─────────────────────────────────────────────────────

TEST_CASE("setChartAxisTitle sets the category axis title",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_axtitle_cat_in",  ".docx");
    const auto out = tmp_file("chart_axtitle_cat_out", ".docx");
    tf_test::write_chart_template_docx(in);   // bar: has catAx/valAx

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartAxisTitle(
        "kStatChart", docweft::IReportMerger::ChartAxis::Category, "Quarter"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto plot_area = doc.child("c:chartSpace").child("c:chart").child("c:plotArea");
    auto cat_title = plot_area.child("c:catAx").child("c:title").child("c:tx")
                     .child("c:rich").child("a:p").child("a:r").child_value("a:t");
    REQUIRE(std::string(cat_title) == "Quarter");
    REQUIRE(!plot_area.child("c:valAx").child("c:title"));

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartAxisTitle sets the value axis title",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_axtitle_val_in",  ".docx");
    const auto out = tmp_file("chart_axtitle_val_out", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartAxisTitle(
        "kStatChart", docweft::IReportMerger::ChartAxis::Value, "USD"));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto plot_area = doc.child("c:chartSpace").child("c:chart").child("c:plotArea");
    auto val_title = plot_area.child("c:valAx").child("c:title").child("c:tx")
                     .child("c:rich").child("a:p").child("a:r").child_value("a:t");
    REQUIRE(std::string(val_title) == "USD");

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartAxisTitle throws InvalidField when the chart has no such axis",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_axtitle_pie", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Pie;   // pie has neither catAx nor valAx
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartAxisTitle(
            "kStatChart", docweft::IReportMerger::ChartAxis::Category, "X");
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartAxisTitle throws NotImplemented on scatter chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_axtitle_scatter", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Scatter;
    tf_test::write_chart_template_docx(in, fx);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartAxisTitle(
            "kStatChart", docweft::IReportMerger::ChartAxis::Value, "X");
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(in);
}

// ── setChartData ───────────────────────────────────────────────────────────

TEST_CASE("setChartData clears a stale per-point <c:dPt> override on reshape",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_dpt_in",  ".docx");
    const auto out = tmp_file("chart_data_dpt_out", ".docx");
    tf_test::ChartFixture fx;
    fx.with_series0_data_point_override = true;   // Sales/idx0 gets a <c:dPt>
    tf_test::write_chart_template_docx(in, fx);

    // Sanity: the fixture really does carry the override before we touch it.
    REQUIRE(read_docx_part(in, "word/charts/chart1.xml").find("c:dPt") != std::string::npos);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart", {"Q1", "Q2", "Q3"},
        {{"Sales", {1, 2, 3}}, {"Costs", {4, 5, 6}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(xml.find("c:dPt") == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData reshapes to more categories than the template had",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_more_in",  ".docx");
    const auto out = tmp_file("chart_data_more_out", ".docx");
    tf_test::write_chart_template_docx(in);   // 2 series × 2 categories

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart",
        {"Q1", "Q2", "Q3", "Q4"},
        {{"Sales", {1, 2, 3, 4}}, {"Costs", {5, 6, 7, 8}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto bar = doc.child("c:chartSpace").child("c:chart").child("c:plotArea").child("c:barChart");

    std::size_t ser_count = 0;
    for (auto ser : bar.children("c:ser")) {
        ++ser_count;
        auto cat = ser.child("c:cat").child("c:strRef");
        REQUIRE(cat.child("c:strCache").child("c:ptCount").attribute("val").as_uint() == 4);
        REQUIRE(std::string(cat.child("c:f").child_value()) == "Sheet1!$A$2:$A$5");
        auto val = ser.child("c:val").child("c:numRef");
        REQUIRE(val.child("c:numCache").child("c:ptCount").attribute("val").as_uint() == 4);
    }
    REQUIRE(ser_count == 2);
    REQUIRE(xml.find(R"(<c:v>Q3</c:v>)") != std::string::npos);
    REQUIRE(xml.find(R"(<c:v>4</c:v>)") != std::string::npos);   // Sales/Q4
    REQUIRE(xml.find(R"(<c:v>8</c:v>)") != std::string::npos);   // Costs/Q4

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData reshapes to fewer categories than the template had",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_fewer_in",  ".docx");
    const auto out = tmp_file("chart_data_fewer_out", ".docx");
    tf_test::write_chart_template_docx(in);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart", {"Only"}, {{"Sales", {42.0}}, {"Costs", {7.0}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto bar = doc.child("c:chartSpace").child("c:chart").child("c:plotArea").child("c:barChart");
    for (auto ser : bar.children("c:ser")) {
        REQUIRE(ser.child("c:cat").child("c:strRef").child("c:strCache")
                   .child("c:ptCount").attribute("val").as_uint() == 1);
    }
    REQUIRE(xml.find("Q1") == std::string::npos);
    REQUIRE(xml.find("Q2") == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData adds a series cloning the template's visual style",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_addser_in",  ".docx");
    const auto out = tmp_file("chart_data_addser_out", ".docx");
    tf_test::write_chart_template_docx(in);   // Sales, Costs

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart", {"Q1", "Q2"},
        {{"Sales", {1, 2}}, {"Costs", {3, 4}}, {"Profit", {5, 6}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto bar = doc.child("c:chartSpace").child("c:chart").child("c:plotArea").child("c:barChart");

    std::vector<std::string> names;
    std::vector<unsigned>    idxs;
    for (auto ser : bar.children("c:ser")) {
        names.emplace_back(ser.child("c:tx").child("c:strRef").child("c:strCache")
                            .child("c:pt").child_value("c:v"));
        idxs.push_back(ser.child("c:idx").attribute("val").as_uint());
    }
    REQUIRE(names == std::vector<std::string>{"Sales", "Costs", "Profit"});
    REQUIRE(idxs == std::vector<unsigned>{0, 1, 2});

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData drops a surplus template series",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_dropser_in",  ".docx");
    const auto out = tmp_file("chart_data_dropser_out", ".docx");
    tf_test::write_chart_template_docx(in);   // Sales, Costs

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart", {"Q1", "Q2"}, {{"OnlyOne", {1, 2}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    const auto xml = read_docx_part(out, "word/charts/chart1.xml");
    pugi::xml_document doc;
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    auto bar = doc.child("c:chartSpace").child("c:chart").child("c:plotArea").child("c:barChart");
    std::size_t ser_count = 0;
    for (auto ser : bar.children("c:ser")) { (void)ser; ++ser_count; }
    REQUIRE(ser_count == 1);
    REQUIRE(xml.find("Costs") == std::string::npos);

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData throws InvalidField on empty categories",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_data_empty_cats", ".docx");
    tf_test::write_chart_template_docx(in);
    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartData("kStatChart", {}, {{"Sales", {}}});
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartData throws InvalidField on empty series",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_data_empty_series", ".docx");
    tf_test::write_chart_template_docx(in);
    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartData("kStatChart", {"Q1"}, {});
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartData throws InvalidField when a series' value count mismatches categories",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_data_mismatch", ".docx");
    tf_test::write_chart_template_docx(in);
    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartData("kStatChart", {"Q1", "Q2"}, {{"Sales", {1.0}}});
        FAIL("expected InvalidField");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidField);
    }
    fs::remove(in);
}

TEST_CASE("setChartData throws NotImplemented on scatter chart",
          "[merger][chart][error]") {
    const auto in = tmp_file("chart_data_scatter", ".docx");
    tf_test::ChartFixture fx;
    fx.kind = tf_test::ChartKind::Scatter;
    tf_test::write_chart_template_docx(in, fx);
    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->setChartData("kStatChart", {"Q1"}, {{"Sales", {1.0}}});
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
    }
    fs::remove(in);
}

TEST_CASE("setChartData before load throws CantOpenTemplate",
          "[merger][chart][error]") {
    auto merger = docweft::make_docx_merger();
    try {
        merger->setChartData("b", {"c"}, {{"s", {1.0}}});
        FAIL("expected CantOpenTemplate");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::CantOpenTemplate);
    }
}

TEST_CASE("setChartData syncs the embedded workbook when the chart has one",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_xlsx_in",  ".docx");
    const auto out = tmp_file("chart_data_xlsx_out", ".docx");
    tf_test::write_chart_template_docx_with_workbook(in);   // Sales/Costs × Q1/Q2

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData(
        "kStatChart", {"Q1", "Q2", "Q3"},
        {{"Sales", {1, 2, 3}}, {"Costs", {4, 5, 6}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    // Chart cache/ranges reflect the new shape.
    const auto chart_xml = read_docx_part(out, "word/charts/chart1.xml");
    REQUIRE(chart_xml.find("Sheet1!$A$2:$A$4") != std::string::npos);

    // The embedded workbook's worksheet was rewritten to match — not left
    // stale, which is the whole point of syncing it in the first place.
    const auto xlsx_bytes = read_docx_part(out, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    REQUIRE(!xlsx_bytes.empty());
    const auto sheet_xml = read_nested_part(xlsx_bytes, "xl/worksheets/sheet1.xml");
    REQUIRE(!sheet_xml.empty());

    pugi::xml_document sheet_doc;
    REQUIRE(sheet_doc.load_buffer(sheet_xml.data(), sheet_xml.size()));
    auto ws = sheet_doc.child("worksheet");
    REQUIRE(std::string(ws.child("dimension").attribute("ref").value()) == "A1:C4");

    // Row 2 = Q1: A2="Q1", B2=1 (Sales), C2=4 (Costs).
    pugi::xml_node row2;
    for (auto row : ws.child("sheetData").children("row")) {
        if (std::string(row.attribute("r").value()) == "2") { row2 = row; break; }
    }
    REQUIRE(row2);
    std::string a2, b2, c2;
    for (auto c : row2.children("c")) {
        const std::string ref = c.attribute("r").value();
        if (ref == "A2") a2 = c.child("is").child_value("t");
        if (ref == "B2") b2 = c.child_value("v");
        if (ref == "C2") c2 = c.child_value("v");
    }
    REQUIRE(a2 == "Q1");
    REQUIRE(b2 == "1");
    REQUIRE(c2 == "4");

    fs::remove(in);
    fs::remove(out);
}

TEST_CASE("setChartData leaves a chart with no embedded workbook alone (no error)",
          "[merger][chart]") {
    const auto in  = tmp_file("chart_data_noxlsx_in",  ".docx");
    const auto out = tmp_file("chart_data_noxlsx_out", ".docx");
    tf_test::write_chart_template_docx(in);   // no <c:externalData> at all

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->setChartData("kStatChart", {"Q1"}, {{"Sales", {1.0}}}));
    REQUIRE_NOTHROW(merger->save(out.string()));

    REQUIRE(read_docx_part(out, "word/embeddings/Microsoft_Excel_Worksheet1.xlsx").empty());

    fs::remove(in);
    fs::remove(out);
}

// ── paste ──────────────────────────────────────────────────────────────────

TEST_CASE("paste records bookmark activation", "[merger][paste]") {
    const auto in = tmp_file("paste_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    REQUIRE_NOTHROW(merger->paste("_Header.User"));

    auto* concrete = dynamic_cast<docweft::docx::DocxMerger*>(merger.get());
    REQUIRE(concrete->pasted_bookmarks().size() == 1);
    REQUIRE(concrete->pasted_bookmarks().front() == "_Header.User");
    fs::remove(in);
}

TEST_CASE("paste throws on unknown bookmark", "[merger][paste][error]") {
    const auto in = tmp_file("paste_err_in", ".docx");
    tf_test::write_minimal_docx(in, kBodyWithHeaderBookmark);

    auto merger = docweft::make_docx_merger();
    merger->load(in.string());
    try {
        merger->paste("_NoSuch");
        FAIL("expected InvalidBookmark");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::InvalidBookmark);
    }
    fs::remove(in);
}

// ── error enum round-trip ──────────────────────────────────────────────────

TEST_CASE("ReportError to_string covers all values", "[error]") {
    using docweft::ReportError;
    using docweft::to_string;
    REQUIRE(to_string(ReportError::None)                 == "None");
    REQUIRE(to_string(ReportError::CantOpenTemplate)     == "CantOpenTemplate");
    REQUIRE(to_string(ReportError::CantCopyDocxTemplate) == "CantCopyDocxTemplate");
    REQUIRE(to_string(ReportError::InvalidBookmark)      == "InvalidBookmark");
    REQUIRE(to_string(ReportError::InvalidField)         == "InvalidField");
    REQUIRE(to_string(ReportError::SaveFailed)           == "SaveFailed");
    REQUIRE(to_string(ReportError::NoConverter)          == "NoConverter");
    REQUIRE(to_string(ReportError::NotImplemented)       == "NotImplemented");
    REQUIRE(to_string(ReportError::Unknown)              == "Unknown");
}
