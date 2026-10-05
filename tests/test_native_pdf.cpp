// Tests for the native PoDoFo PDF renderer (docx/pdf_native.cpp): content
// that must not be dropped, and strict mode rejecting what it can't draw.
// Only built into the test binary when DOCWEFT_ENABLE_NATIVE_PDF is on.

#if defined(DOCWEFT_HAVE_PODOFO)

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include <podofo/podofo.h>
#include <pugixml.hpp>

#include "docweft/error.hpp"
#include "docx/merger.hpp"
#include "docx/pdf_native.hpp"
#include "fixture_helpers.hpp"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <set>
#include <sstream>
#include <cstdlib>
#include <filesystem>
#include <random>
#include <string>
#include <unordered_map>
#include <vector>

namespace fs = std::filesystem;

namespace {

using Parts = std::unordered_map<std::string, std::string>;

// A temp PDF path no other test process uses (ctest -j runs each case in
// its own process, so fixed names race).
fs::path temp_pdf(const std::string& stem) {
    static const unsigned long long tag = std::random_device{}() ^
        static_cast<unsigned long long>(std::chrono::steady_clock::now().time_since_epoch().count());
    static unsigned counter = 0;
    return fs::temp_directory_path() /
           (stem + "_" + std::to_string(tag) + "_" + std::to_string(counter++) + ".pdf");
}

constexpr const char* kNamespaces =
    R"( xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main")"
    R"( xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships")"
    R"( xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")"
    R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main")"
    R"( xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture")"
    R"( xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart")";

constexpr const char* kRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

std::string p(const std::string& text, const std::string& ppr = {}) {
    return "<w:p><w:pPr>" + ppr + "</w:pPr><w:r><w:t xml:space=\"preserve\">" + text +
           "</w:t></w:r></w:p>";
}

struct Rel {
    std::string id, type, target;
};

// A package with `body` (plus `sect_pr`) as word/document.xml and any extra parts.
Parts make_parts(const std::string& body, const std::vector<Rel>& rels = {},
                 const Parts& extra = {}, const std::string& sect_pr = {}) {
    Parts parts = extra;
    parts["word/document.xml"] =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document)") +
        kNamespaces + "><w:body>" + body + sect_pr + "</w:body></w:document>";
    std::string r = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
                    R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";
    for (const Rel& rel : rels) {
        r += "<Relationship Id=\"" + rel.id + "\" Type=\"" + kRelNs + rel.type +
             "\" Target=\"" + rel.target + "\"/>";
    }
    r += "</Relationships>";
    parts["word/_rels/document.xml.rels"] = r;
    return parts;
}

// Renders `parts` and returns the text of each page.
std::vector<std::string> render_pages(const Parts& parts) {
    pugi::xml_document doc;
    const std::string& xml = parts.at("word/document.xml");
    REQUIRE(doc.load_buffer(xml.data(), xml.size(), docweft::docx::kXmlParse));

    const fs::path out = temp_pdf("docweft_native_test");
    docweft::docx::render_native_pdf(doc, parts, out);

    PoDoFo::PdfMemDocument pdf;
    pdf.Load(out.string());
    std::vector<std::string> pages;
    for (unsigned i = 0; i < pdf.GetPages().GetCount(); ++i) {
        std::vector<PoDoFo::PdfTextEntry> entries;
        pdf.GetPages().GetPageAt(i).ExtractTextTo(entries);
        std::string text;
        for (const auto& e : entries) text += e.Text + "\n";
        pages.push_back(text);
    }
    fs::remove(out);
    return pages;
}

struct TextEntry {
    std::string text;
    int         page = 0;
    double      x = 0.0, y = 0.0;
};

// Renders `parts` and returns every text entry with its position.
std::vector<TextEntry> render_entries(const Parts& parts) {
    pugi::xml_document doc;
    const std::string& xml = parts.at("word/document.xml");
    REQUIRE(doc.load_buffer(xml.data(), xml.size(), docweft::docx::kXmlParse));

    const fs::path out = temp_pdf("docweft_native_test_entries");
    docweft::docx::render_native_pdf(doc, parts, out);

    PoDoFo::PdfMemDocument pdf;
    pdf.Load(out.string());
    std::vector<TextEntry> result;
    for (unsigned i = 0; i < pdf.GetPages().GetCount(); ++i) {
        std::vector<PoDoFo::PdfTextEntry> entries;
        pdf.GetPages().GetPageAt(i).ExtractTextTo(entries);
        for (const auto& e : entries) result.push_back({e.Text, static_cast<int>(i), e.X, e.Y});
    }
    fs::remove(out);
    return result;
}

const TextEntry& entry_with(const std::vector<TextEntry>& entries, const std::string& needle) {
    for (const auto& e : entries) {
        if (e.text.find(needle) != std::string::npos) return e;
    }
    FAIL("no text entry containing " << needle);
    static TextEntry none;
    return none;
}

std::string all_text(const std::vector<std::string>& pages) {
    std::string out;
    for (const auto& page : pages) out += page;
    return out;
}

bool contains(const std::string& haystack, const std::string& needle) {
    return haystack.find(needle) != std::string::npos;
}

// RAII env var override (restores the previous value).
struct ScopedEnvVar {
    const char* name;
    std::string prev;
    bool        had_prev = false;
    ScopedEnvVar(const char* n, const char* value) : name(n) {
        if (const char* v = std::getenv(n)) { had_prev = true; prev = v; }
        set(value);
    }
    ~ScopedEnvVar() { set(had_prev ? prev.c_str() : nullptr); }
    void set(const char* value) {
#ifdef _WIN32
        _putenv_s(name, value ? value : "");
#else
        if (value) ::setenv(name, value, 1);
        else       ::unsetenv(name);
#endif
    }
};

std::string chart_xml(const std::string& type_element_open, const std::string& type_element_close,
                      const std::string& title = {}) {
    std::string title_xml;
    if (!title.empty()) {
        title_xml = "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>" + title +
                    "</a:t></a:r></a:p></c:rich></c:tx></c:title>";
    }
    return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") +
           kNamespaces + "><c:chart>" + title_xml + "<c:plotArea>" + type_element_open +
           R"(<c:ser><c:idx val="0"/><c:order val="0"/>)"
           R"(<c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series</c:v></c:pt></c:strCache></c:strRef></c:tx>)"
           R"(<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>)"
           R"(<c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>)"
           "</c:ser>" + type_element_close +
           R"(<c:catAx><c:axId val="1"/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>)"
           "</c:plotArea></c:chart></c:chartSpace>";
}

// A <wps:wsp> drawing in a run: inline, or anchored in front of the text
// with no wrapping at (x, y) points from the page's top-left corner.
// `sp_pr` goes into <wps:spPr> after the geometry, `style` is the whole
// <wps:style> (or empty), `body_pr` the <wps:bodyPr> attributes and
// children, `content` the text box's blocks (no text box if empty).
std::string shape_run(bool anchored, double x, double y, double w, double h, const std::string& geom,
                      const std::string& sp_pr, const std::string& style = {},
                      const std::string& content = {}, const std::string& body_pr = "<wps:bodyPr/>") {
    auto emu = [](double pt) { return std::to_string(static_cast<long long>(pt * 12700)); };
    const std::string extent = R"(<wp:extent cx=")" + emu(w) + R"(" cy=")" + emu(h) + R"("/>)";
    const std::string graphic =
        R"(<wp:docPr id="7" name="s"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">)"
        R"(<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr>)"
        R"(<a:xfrm><a:off x="0" y="0"/><a:ext cx=")" + emu(w) + R"(" cy=")" + emu(h) + R"("/></a:xfrm>)"
        R"(<a:prstGeom prst=")" + geom + R"("><a:avLst/></a:prstGeom>)" + sp_pr + "</wps:spPr>" + style +
        (content.empty() ? "" : "<wps:txbx><w:txbxContent>" + content + "</w:txbxContent></wps:txbx>") + body_pr +
        "</wps:wsp></a:graphicData></a:graphic>";
    if (!anchored) return "<w:r><w:drawing><wp:inline>" + extent + graphic + "</wp:inline></w:drawing></w:r>";
    return R"(<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">)"
           R"(<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>)" + emu(x) +
           R"(</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>)" + emu(y) +
           "</wp:posOffset></wp:positionV>" + extent + "<wp:wrapNone/>" + graphic + "</wp:anchor></w:drawing></w:r>";
}

std::string chart_paragraph(const std::string& rid) {
    return R"(<w:p><w:r><w:drawing><wp:inline><wp:extent cx="3000000" cy="1800000"/>)"
           R"(<wp:docPr id="1" name="c"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">)"
           R"(<c:chart r:id=")" + rid + R"("/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>)";
}

}  // namespace

TEST_CASE("native PDF keeps text inside hyperlinks, content controls and fields",
          "[native_pdf]") {
    const std::string body =
        R"(<w:p><w:r><w:t xml:space="preserve">Link: </w:t></w:r>)"
        R"(<w:hyperlink r:id="rIdL"><w:r><w:t>HYPERLINK-TEXT</w:t></w:r></w:hyperlink></w:p>)"
        R"(<w:sdt><w:sdtContent>)" + p("BLOCK-SDT-TEXT") + R"(</w:sdtContent></w:sdt>)"
        R"(<w:p><w:sdt><w:sdtContent><w:r><w:t>INLINE-SDT-TEXT</w:t></w:r></w:sdtContent></w:sdt>)"
        R"(<w:smartTag><w:r><w:t xml:space="preserve"> SMARTTAG-TEXT</w:t></w:r></w:smartTag>)"
        R"(<w:ins><w:r><w:t xml:space="preserve"> INSERTED-TEXT</w:t></w:r></w:ins>)"
        R"(<w:del><w:r><w:delText>DELETED-TEXT</w:delText></w:r></w:del>)"
        R"(<w:fldSimple w:instr=" DATE "><w:r><w:t xml:space="preserve"> CACHED-DATE</w:t></w:r></w:fldSimple>)"
        R"(<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> AUTHOR </w:instrText></w:r>)"
        R"(<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t xml:space="preserve"> CACHED-AUTHOR</w:t></w:r>)"
        R"(<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>)";
    const std::string text = all_text(render_pages(make_parts(body)));

    for (const char* expected : {"HYPERLINK-TEXT", "BLOCK-SDT-TEXT", "INLINE-SDT-TEXT", "SMARTTAG-TEXT",
                                 "INSERTED-TEXT", "CACHED-DATE", "CACHED-AUTHOR"}) {
        INFO(expected);
        REQUIRE(contains(text, expected));
    }
    REQUIRE_FALSE(contains(text, "DELETED-TEXT"));
    REQUIRE_FALSE(contains(text, "AUTHOR "));  // the field instruction itself
}

TEST_CASE("native PDF keeps text in content-control-wrapped table rows and cells",
          "[native_pdf]") {
    const std::string body =
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>)"
        R"(<w:sdt><w:sdtContent><w:tr><w:tc>)" + p("ROW-SDT-A") + "</w:tc><w:tc>" + p("ROW-SDT-B") +
        R"(</w:tc></w:tr></w:sdtContent></w:sdt>)"
        R"(<w:tr><w:sdt><w:sdtContent><w:tc>)" + p("CELL-SDT") + R"(</w:tc></w:sdtContent></w:sdt>)"
        "<w:tc>" + p("CELL-TEXT-BEFORE") +
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>)" + p("NESTED-TEXT") +
        "</w:tc></w:tr></w:tbl>" + p("CELL-TEXT-AFTER") + "</w:tc></w:tr></w:tbl>";
    const std::string text = all_text(render_pages(make_parts(body)));

    for (const char* expected : {"ROW-SDT-A", "ROW-SDT-B", "CELL-SDT", "CELL-TEXT-BEFORE",
                                 "NESTED-TEXT", "CELL-TEXT-AFTER"}) {
        INFO(expected);
        REQUIRE(contains(text, expected));
    }
}

TEST_CASE("native PDF honours manual page breaks", "[native_pdf]") {
    const std::string body = p("PAGE-ONE") +
        R"(<w:p><w:r><w:t>STILL-ONE</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>PAGE-TWO</w:t></w:r></w:p>)" +
        p("PAGE-THREE", "<w:pageBreakBefore/>");
    const auto pages = render_pages(make_parts(body));

    REQUIRE(pages.size() == 3);
    REQUIRE(contains(pages[0], "PAGE-ONE"));
    REQUIRE(contains(pages[0], "STILL-ONE"));
    REQUIRE(contains(pages[1], "PAGE-TWO"));
    REQUIRE(contains(pages[2], "PAGE-THREE"));
}

TEST_CASE("native PDF draws headers and footers with page numbers", "[native_pdf]") {
    auto hf = [](const char* root, const std::string& content) {
        return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:)") + root +
               kNamespaces + ">" + content + "</w:" + root + ">";
    };
    const std::string footer = hf("ftr",
        R"(<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>)"
        R"(<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>)"
        R"(<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>)"
        R"(<w:r><w:t xml:space="preserve"> of </w:t></w:r>)"
        R"(<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>)");
    const std::string first_header = hf("hdr", p("FIRST-PAGE-HEADER"));
    const std::string header       = hf("hdr", p("DEFAULT-HEADER"));

    const std::string body = p("BODY-ONE") + R"(<w:p><w:r><w:br w:type="page"/></w:r></w:p>)" + p("BODY-TWO");
    const std::string sect =
        R"(<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>)"
        R"(<w:headerReference w:type="first" r:id="rIdH1"/>)"
        R"(<w:footerReference w:type="default" r:id="rIdF"/><w:titlePg/></w:sectPr>)";
    const auto pages = render_pages(make_parts(
        body,
        {{"rIdH", "header", "header1.xml"}, {"rIdH1", "header", "header2.xml"},
         {"rIdF", "footer", "footer1.xml"}},
        {{"word/header1.xml", header}, {"word/header2.xml", first_header},
         {"word/footer1.xml", footer}},
        sect));

    REQUIRE(pages.size() == 2);
    REQUIRE(contains(pages[0], "FIRST-PAGE-HEADER"));
    REQUIRE_FALSE(contains(pages[0], "DEFAULT-HEADER"));
    REQUIRE(contains(pages[1], "DEFAULT-HEADER"));
    // titlePg with no first-page footer: page 1 has none.
    REQUIRE_FALSE(contains(pages[0], "Page 1 of 2"));
    REQUIRE(contains(pages[1], "Page 2 of 2"));
}

TEST_CASE("native PDF draws list markers", "[native_pdf]") {
    const std::string numbering =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering)") + kNamespaces + ">"
        R"(<w:abstractNum w:abstractNumId="0">)"
        R"(<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>)"
        R"x(<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="russianLower"/><w:lvlText w:val="%1.%2)"/></w:lvl>)x"
        R"(</w:abstractNum>)"
        R"(<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/></w:lvl></w:abstractNum>)"
        R"(<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>)"
        R"(<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>)";
    auto item = [](const char* text, int num, int lvl) {
        return p(text, "<w:numPr><w:ilvl w:val=\"" + std::to_string(lvl) + "\"/><w:numId w:val=\"" +
                           std::to_string(num) + "\"/></w:numPr>");
    };
    const std::string body = item("first", 1, 0) + item("second", 1, 0) + item("sub-a", 1, 1) +
                             item("sub-b", 1, 1) + item("third", 1, 0) + item("bullet", 2, 0);
    const std::string text = all_text(render_pages(
        make_parts(body, {{"rIdN", "numbering", "numbering.xml"}}, {{"word/numbering.xml", numbering}})));

    for (const char* expected : {"1.", "2.", "2.а)", "2.б)", "3.", "•"}) {
        INFO(expected);
        REQUIRE(contains(text, expected));
    }
}

TEST_CASE("native PDF draws Symbol and Wingdings characters as Unicode", "[native_pdf]") {
    const std::string numbering =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering)") + kNamespaces + ">"
        R"(<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0D8;"/>)"
        R"(<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl></w:abstractNum>)"
        R"(<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>)";
    const std::string body =
        R"(<w:p><w:r><w:t xml:space="preserve">done </w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r>)"
        R"(<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r><w:r><w:sym w:font="Symbol" w:char="F061"/></w:r>)"
        R"(<w:r><w:sym w:font="Symbol" w:char="B3"/></w:r></w:p>)"
        R"(<w:p><w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:t>&#xFD;o</w:t></w:r></w:p>)"
        R"(<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>)";
    // Strict mode doesn't object: every symbol here has a Unicode equivalent.
    const std::string text = all_text(render_pages(
        make_parts(body, {{"rIdN", "numbering", "numbering.xml"}}, {{"word/numbering.xml", numbering}})));
    for (const char* expected : {"☑", "✔", "α", "≥", "☒", "□", "➢"}) {
        INFO(expected);
        REQUIRE(contains(text, expected));
    }
    REQUIRE_FALSE(contains(text, "\xC3\xBD"));  // "ý", the Wingdings byte itself
}

TEST_CASE("native PDF keeps a run that holds only a space", "[native_pdf]") {
    const auto entries = render_entries(make_parts(
        R"(<w:p><w:r><w:t>AAA</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:t>BBB</w:t></w:r></w:p>)"));
    std::string line;
    for (const auto& e : entries) line += "[" + e.text + "@" + std::to_string(e.x) + "]";
    INFO(line);
    CHECK(contains(line, "AAA BBB"));
}


TEST_CASE("native PDF draws chart and axis titles", "[native_pdf]") {
    const auto parts = make_parts(
        chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
        {{"word/charts/chart1.xml",
          chart_xml(R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)", "</c:barChart>",
                    "SALES-TITLE")}});
    const std::string text = all_text(render_pages(parts));
    REQUIRE(contains(text, "SALES-TITLE"));
    REQUIRE(contains(text, "Quarter"));
}

TEST_CASE("native PDF value axis uses round steps, not raw data bounds", "[native_pdf]") {
    std::string chart = chart_xml(R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)",
                                  "</c:barChart>");
    // Values 12450 and 21080: expect 0..25000 by 5000, like Excel/LibreOffice.
    const auto replace = [&](const std::string& from, const std::string& to) {
        chart.replace(chart.find(from), from.size(), to);
    };
    replace("<c:v>3</c:v>", "<c:v>12450</c:v>");
    replace("<c:v>5</c:v>", "<c:v>21080</c:v>");
    const std::string text = all_text(render_pages(make_parts(
        chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
        {{"word/charts/chart1.xml", chart}})));

    for (const char* label : {"0\n", "5000", "10000", "15000", "20000", "25000"}) {
        INFO(label);
        REQUIRE(contains(text, label));
    }
    REQUIRE_FALSE(contains(text, "e+"));
    REQUIRE_FALSE(contains(text, "21080"));  // the raw maximum is not a tick
}

// A4, 2 cm margins: text area x = 56.7 .. 538.6 pt.
constexpr const char* kSect =
    R"(<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>)";

TEST_CASE("native PDF aligns paragraphs and honours right tab stops", "[native_pdf]") {
    const std::string body =
        p("LEFT-TEXT") + p("CENTER-TEXT", R"(<w:jc w:val="center"/>)") + p("RIGHT-TEXT", R"(<w:jc w:val="right"/>)") +
        p("INDENTED-TEXT", R"(<w:ind w:left="1134"/>)") +
        R"(<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="9638"/></w:tabs></w:pPr>)"
        R"(<w:r><w:t>Company</w:t></w:r><w:r><w:tab/><w:t>TABBED-RIGHT</w:t></w:r></w:p>)";
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));

    const double left   = entry_with(entries, "LEFT-TEXT").x;
    const double center = entry_with(entries, "CENTER-TEXT").x;
    const double right  = entry_with(entries, "RIGHT-TEXT").x;
    REQUIRE(left == Catch::Approx(56.7).margin(1.0));
    REQUIRE(center > left + 100.0);
    REQUIRE(right > center + 100.0);
    REQUIRE(entry_with(entries, "INDENTED-TEXT").x == Catch::Approx(56.7 + 56.7).margin(1.0));
    // Right tab at 17 cm from the margin: the text ends there, well right of centre.
    REQUIRE(entry_with(entries, "TABBED-RIGHT").x > 400.0);
}

TEST_CASE("native PDF applies paragraph spacing", "[native_pdf]") {
    const std::string body = p("FIRST-PARA", R"(<w:spacing w:after="720"/>)") + p("SECOND-PARA") +
                             p("THIRD-PARA");
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));
    const double gap_after_spacing = entry_with(entries, "FIRST-PARA").y - entry_with(entries, "SECOND-PARA").y;
    const double plain_gap         = entry_with(entries, "SECOND-PARA").y - entry_with(entries, "THIRD-PARA").y;
    REQUIRE(gap_after_spacing == Catch::Approx(plain_gap + 36.0).margin(1.0));  // 720 twips = 36 pt
}

TEST_CASE("native PDF repeats table header rows on every page", "[native_pdf]") {
    std::string rows;
    for (int i = 0; i < 120; ++i) rows += "<w:tr><w:tc>" + p("row " + std::to_string(i)) + "</w:tc></w:tr>";
    const std::string body =
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>)"
        R"(<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr>)" +
        p("HEADER-ROW") + "</w:tc></w:tr>" + rows + "</w:tbl>";
    const auto pages = render_pages(make_parts(body, {}, {}, kSect));

    REQUIRE(pages.size() >= 2);
    for (const auto& page : pages) REQUIRE(contains(page, "HEADER-ROW"));
}

// A chart part with the given plot-area content (groups and axes).
// `after`: what follows the plot area in <c:chart> (a <c:legend>).
std::string chart_with_plot(const std::string& plot, const std::string& after = {}) {
    return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") + kNamespaces +
           "><c:chart><c:plotArea>" + plot + "</c:plotArea>" + after + "</c:chart></c:chartSpace>";
}

std::string series_xml(int idx, const std::string& name, double q1, double q2) {
    auto num = [](double v) { std::ostringstream os; os << v; return os.str(); };
    return "<c:ser><c:idx val=\"" + std::to_string(idx) + "\"/><c:order val=\"" + std::to_string(idx) + "\"/>"
           "<c:tx><c:strRef><c:strCache><c:pt idx=\"0\"><c:v>" + name + "</c:v></c:pt></c:strCache></c:strRef></c:tx>"
           R"(<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>)"
           R"(<c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>)" + num(q1) +
           R"(</c:v></c:pt><c:pt idx="1"><c:v>)" + num(q2) + "</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>";
}

constexpr const char* kAxes =
    R"(<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>)"
    R"(<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>)";

std::vector<TextEntry> render_chart(const std::string& plot, const std::string& after = {}) {
    return render_entries(make_parts(chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
                                     {{"word/charts/chart1.xml", chart_with_plot(plot, after)}}, kSect));
}

bool has_entry(const std::vector<TextEntry>& entries, const std::string& text) {
    return std::any_of(entries.begin(), entries.end(), [&](const TextEntry& e) { return e.text == text; });
}

TEST_CASE("native PDF draws vertically merged cells once, spanning their rows", "[native_pdf]") {
    std::string merged;
    for (int i = 1; i <= 6; ++i) merged += p("M" + std::to_string(i));
    const std::string body =
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>)"
        R"(<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>)" + merged + "</w:tc><w:tc>" + p("R1") + "</w:tc></w:tr>"
        R"(<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc>)" + p("R2") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p("C3") + "</w:tc><w:tc>" + p("R3") + "</w:tc></w:tr></w:tbl>";
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));

    REQUIRE(std::count_if(entries.begin(), entries.end(), [](const TextEntry& e) { return e.text == "M1"; }) == 1);
    // Row 2 starts right under row 1 (next to the merged cell's text) …
    REQUIRE(entry_with(entries, "R2").y > entry_with(entries, "M6").y);
    // … and grows so row 3 starts below the whole merged cell.
    REQUIRE(entry_with(entries, "R3").y < entry_with(entries, "M6").y);
}

TEST_CASE("native PDF draws horizontal and stacked bar charts", "[native_pdf]") {
    SECTION("horizontal: categories along the left edge, first at the bottom") {
        const auto entries = render_chart(
            R"(<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>)" + series_xml(0, "A", 3, 5) +
            R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes);
        const TextEntry& q1 = entry_with(entries, "Q1");
        const TextEntry& q2 = entry_with(entries, "Q2");
        REQUIRE(q1.y < q2.y);
        REQUIRE(q1.x == Catch::Approx(q2.x).margin(3.0));
    }
    SECTION("stacked: the value axis covers the sums") {
        const auto entries = render_chart(
            R"(<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>)" + series_xml(0, "A", 12450, 21080) +
            series_xml(1, "B", 10000, 10000) + R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes);
        // Sum 31080 (+5 % headroom): the axis goes past it; clustered would stop at 25000.
        REQUIRE((has_entry(entries, "35000") || has_entry(entries, "40000")));
    }
    SECTION("100 % stacked: a percent axis") {
        const auto entries = render_chart(
            R"(<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/>)" + series_xml(0, "A", 1, 3) +
            series_xml(1, "B", 3, 1) + R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes);
        REQUIRE(has_entry(entries, "100%"));
        REQUIRE(has_entry(entries, "0%"));
    }
}

TEST_CASE("native PDF draws combo charts with a secondary axis", "[native_pdf]") {
    const auto entries = render_chart(
        R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)" + series_xml(0, "Revenue", 12450, 21080) +
        R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)"
        R"(<c:lineChart><c:grouping val="standard"/>)" + series_xml(1, "Margin", 22, 27) +
        R"(<c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>)" + kAxes +
        R"(<c:catAx><c:axId val="3"/><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>)"
        R"(<c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/></c:valAx>)",
        R"(<c:legend><c:legendPos val="b"/></c:legend>)");
    REQUIRE(has_entry(entries, "25000"));  // primary axis, left
    REQUIRE(has_entry(entries, "30"));     // secondary axis, right
    REQUIRE(entry_with(entries, "30").x > entry_with(entries, "25000").x + 100.0);
    REQUIRE(has_entry(entries, "Revenue"));
    REQUIRE(has_entry(entries, "Margin"));
}

TEST_CASE("native PDF puts the chart legend where <c:legendPos> says", "[native_pdf]") {
    const std::string bars = R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)" +
                             series_xml(0, "Alpha", 3, 5) + series_xml(1, "Beta", 4, 2) +
                             R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes;
    SECTION("no <c:legend>: no legend") {
        const auto entries = render_chart(bars);
        CHECK_FALSE(has_entry(entries, "Alpha"));
        CHECK(has_entry(entries, "Q1"));
    }
    SECTION("bottom: one centred row under the category labels") {
        const auto entries = render_chart(bars, R"(<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>)");
        const TextEntry& a = entry_with(entries, "Alpha");
        const TextEntry& b = entry_with(entries, "Beta");
        CHECK(a.y == Catch::Approx(b.y).margin(0.5));
        CHECK(a.y < entry_with(entries, "Q1").y);
        // One row: extracted as one piece of text, or two side by side.
        if (&a == &b) CHECK(a.text.find("Alpha") < a.text.find("Beta"));
        else CHECK(b.x > a.x);
    }
    SECTION("right (also the default position): a column right of the plot, inside the chart") {
        for (const char* legend : {R"(<c:legend><c:legendPos val="r"/></c:legend>)", "<c:legend/>"}) {
            INFO(legend);
            const auto entries = render_chart(bars, legend);
            const TextEntry& a = entry_with(entries, "Alpha");
            const TextEntry& b = entry_with(entries, "Beta");
            CHECK(a.x == Catch::Approx(b.x).margin(0.5));
            CHECK(a.y > b.y);
            CHECK(a.x > entry_with(entries, "Q2").x);
            // Within the chart's 236 pt wide extent from the left margin.
            CHECK(a.x < 56.7 + 236.0);
            CHECK(b.y > entry_with(entries, "Q1").y);
        }
    }
    SECTION("a deleted entry is left out") {
        const auto entries = render_chart(
            bars, R"(<c:legend><c:legendPos val="t"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry></c:legend>)");
        CHECK(has_entry(entries, "Alpha"));
        CHECK_FALSE(has_entry(entries, "Beta"));
        CHECK(entry_with(entries, "Alpha").y > entry_with(entries, "Q1").y);
    }
    SECTION("pie: the categories, right of the pie") {
        const auto entries = render_chart(R"(<c:pieChart><c:varyColors val="1"/>)" + series_xml(0, "S", 3, 5) +
                                              "</c:pieChart>",
                                          R"(<c:legend><c:legendPos val="r"/></c:legend>)");
        const TextEntry& q1 = entry_with(entries, "Q1");
        const TextEntry& q2 = entry_with(entries, "Q2");
        CHECK(q1.x == Catch::Approx(q2.x).margin(0.5));
        CHECK(q1.y > q2.y);
        CHECK(q1.x > 56.7 + 236.0 / 2.0);  // past the middle of the chart, where the pie is
    }
}

TEST_CASE("native PDF lays out several sections", "[native_pdf]") {
    auto hf = [](const char* root, const std::string& content) {
        return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:)") + root +
               kNamespaces + ">" + content + "</w:" + root + ">";
    };
    const std::string footer = hf("ftr",
        R"(<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>)");
    const std::string body =
        p("PORTRAIT-BODY") +
        R"(<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF"/>)"
        R"(<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:pPr></w:p>)" +
        p("LANDSCAPE-BODY");
    const std::string last_sect =
        R"(<w:sectPr><w:headerReference w:type="default" r:id="rIdH2"/>)"
        R"(<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>)"
        R"(<w:pgNumType w:start="7"/></w:sectPr>)";
    const Parts parts = make_parts(
        body,
        {{"rIdH1", "header", "header1.xml"}, {"rIdH2", "header", "header2.xml"}, {"rIdF", "footer", "footer1.xml"}},
        {{"word/header1.xml", hf("hdr", p("FIRST-SECTION-HEADER"))},
         {"word/header2.xml", hf("hdr", p("SECOND-SECTION-HEADER"))},
         {"word/footer1.xml", footer}},
        last_sect);

    pugi::xml_document doc;
    REQUIRE(doc.load_string(parts.at("word/document.xml").c_str()));
    const fs::path out = temp_pdf("docweft_native_sections");
    docweft::docx::render_native_pdf(doc, parts, out);
    PoDoFo::PdfMemDocument pdf;
    pdf.Load(out.string());
    REQUIRE(pdf.GetPages().GetCount() == 2);
    const auto r1 = pdf.GetPages().GetPageAt(0).GetRect();
    const auto r2 = pdf.GetPages().GetPageAt(1).GetRect();
    REQUIRE(r1.Height > r1.Width);
    REQUIRE(r2.Width > r2.Height);
    fs::remove(out);

    const auto pages = render_pages(parts);
    REQUIRE(contains(pages[0], "FIRST-SECTION-HEADER"));
    REQUIRE(contains(pages[0], "Page 1"));
    REQUIRE(contains(pages[1], "SECOND-SECTION-HEADER"));  // its own header
    REQUIRE(contains(pages[1], "Page 7"));                 // footer inherited, numbering restarted
    REQUIRE(contains(pages[1], "LANDSCAPE-BODY"));
}

TEST_CASE("native PDF keeps a keepNext paragraph with the next one", "[native_pdf]") {
    // Exact 12 pt lines make the page arithmetic font-independent: the body
    // is 728.5 pt, 58 filler lines leave 32.5 pt — room for the heading
    // (12 pt) but not for it plus the 30 pt line after it.
    const std::string exact = R"(<w:spacing w:line="240" w:lineRule="exact"/>)";
    std::string filler;
    for (int i = 1; i <= 58; ++i) filler += p("FILLER-" + std::to_string(i), exact);
    const std::string tall_next = p("NEXT-BLOCK", R"(<w:spacing w:line="600" w:lineRule="exact"/>)");

    SECTION("with keepNext the heading moves to the next page") {
        const auto pages = render_pages(make_parts(
            filler + p("HEADING", exact + "<w:keepNext/>") + tall_next, {}, {}, kSect));
        REQUIRE(contains(pages[0], "FILLER-58"));
        REQUIRE_FALSE(contains(pages[0], "HEADING"));
        REQUIRE(contains(pages[1], "HEADING"));
    }
    SECTION("without it the heading stays at the bottom") {
        const auto pages = render_pages(make_parts(filler + p("HEADING", exact) + tall_next, {}, {}, kSect));
        REQUIRE(contains(pages[0], "HEADING"));
    }
}

TEST_CASE("native PDF strict mode rejects content it would get wrong", "[native_pdf]") {
    const std::string symbol =
        R"(<w:p><w:r><w:t>With a symbol</w:t></w:r><w:r><w:sym w:font="Webdings" w:char="F021"/></w:r></w:p>)";
    const std::string columns =
        R"(<w:sectPr><w:cols w:num="2"/></w:sectPr>)";

    SECTION("strict (default): NotImplemented listing every problem") {
        const auto parts = make_parts(
            symbol + R"(<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>)" + chart_paragraph("rIdC"),
            {{"rIdC", "chart", "charts/chart1.xml"}},
            {{"word/charts/chart1.xml",
              chart_with_plot(R"(<c:pieChart>)" + series_xml(0, "A", 1, 2) + R"(</c:pieChart>)"
                              R"(<c:barChart><c:barDir val="col"/>)" + series_xml(1, "B", 1, 2) + R"(</c:barChart>)")}},
            columns);
        try {
            render_pages(parts);
            FAIL("expected NotImplemented");
        } catch (const docweft::ReportException& e) {
            REQUIRE(e.code() == docweft::ReportError::NotImplemented);
            const std::string msg = e.what();
            REQUIRE(contains(msg, "symbol character F021 of font 'Webdings'"));
            REQUIRE(contains(msg, "footnotes in multi-column layout"));
            REQUIRE(contains(msg, "pie charts combined with other chart types"));
            REQUIRE(contains(msg, "DOCWEFT_NATIVE_PDF_STRICT"));
        }
    }
    SECTION("DOCWEFT_NATIVE_PDF_STRICT=0 renders anyway") {
        ScopedEnvVar lenient("DOCWEFT_NATIVE_PDF_STRICT", "0");
        REQUIRE(contains(all_text(render_pages(make_parts(symbol, {}, {}, columns))), "With a symbol"));
    }
}

TEST_CASE("native PDF strict mode rejects shapes it can't draw and unsupported images", "[native_pdf]") {
    const std::string group =
        R"(<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="1000" cy="1000"/><a:graphic>)"
        R"(<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">)"
        R"(</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>)";
    const std::string star = "<w:p>" + shape_run(false, 0, 0, 50, 50, "star5", "") + "</w:p>";
    const std::string rotated = "<w:p>" + shape_run(false, 0, 0, 50, 50, "rect", "") + "</w:p>";
    const std::string float_in_box =
        "<w:p>" + shape_run(true, 0, 0, 100, 50, "rect", "", {}, "<w:p>" + shape_run(true, 10, 10, 5, 5, "rect", "") + "</w:p>") + "</w:p>";
    const std::string emf_picture =
        R"(<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1000" cy="1000"/><a:graphic>)"
        R"(<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill>)"
        R"(<a:blip r:embed="rIdE"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>)";
    std::string emf(44, '\0');
    emf.replace(40, 4, " EMF");

    try {
        std::string turned = rotated;
        turned.replace(turned.find("<a:xfrm>"), 8, R"(<a:xfrm rot="2700000">)");
        render_pages(make_parts(group + star + turned + float_in_box + emf_picture, {{"rIdE", "image", "media/image1.emf"}},
                                {{"word/media/image1.emf", emf}}));
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
        const std::string msg = e.what();
        CHECK(contains(msg, "shape groups"));
        CHECK(contains(msg, "shapes of type 'star5'"));
        CHECK(contains(msg, "rotated shapes"));
        CHECK(contains(msg, "floating objects inside text boxes"));
        CHECK(contains(msg, "EMF"));
    }
}

TEST_CASE("native PDF keeps a paragraph's text next to its picture", "[native_pdf]") {
    const std::string picture =
        R"(<w:p><w:r><w:t xml:space="preserve">CAPTION-BEFORE </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>)"
        R"(<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill>)"
        R"(<a:blip r:embed="rIdP"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>)"
        R"(<w:r><w:t xml:space="preserve"> CAPTION-AFTER</w:t></w:r></w:p>)";
    const std::string text = all_text(render_pages(make_parts(
        picture, {{"rIdP", "image", "media/image1.png"}},
        {{"word/media/image1.png", std::string(tf_test::tiny_png_bytes())}})));
    REQUIRE(contains(text, "CAPTION-BEFORE"));
    REQUIRE(contains(text, "CAPTION-AFTER"));
}

namespace {

std::string r(const std::string& text, const std::string& rpr = {}) {
    return "<w:r><w:rPr>" + rpr + "</w:rPr><w:t xml:space=\"preserve\">" + text + "</w:t></w:r>";
}

const std::string kHiddenStyles =
    std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
    R"(<w:style w:type="character" w:styleId="Hid"><w:rPr><w:vanish/></w:rPr></w:style>)"
    R"(<w:style w:type="paragraph" w:styleId="HidPara"><w:rPr><w:vanish/></w:rPr></w:style>)"
    "</w:styles>";

Parts with_hidden_styles(const std::string& body, const std::vector<Rel>& rels = {}, Parts extra = {}) {
    std::vector<Rel> all = rels;
    all.push_back({"rIdS", "styles", "styles.xml"});
    extra["word/styles.xml"] = kHiddenStyles;
    return make_parts(body, all, extra, kSect);
}

}  // namespace

TEST_CASE("native PDF leaves out hidden text but prints web-hidden text", "[native_pdf]") {
    const std::string body =
        "<w:p>" + r("SHOWN-A ") + r("HIDDEN-DIRECT ", "<w:vanish/>") + r("SHOWN-B") + "</w:p>" +
        "<w:p>" + r("HIDDEN-BY-CHAR-STYLE", R"(<w:rStyle w:val="Hid"/>)") + "</w:p>" +
        "<w:p>" + r("SHOWN-OVERRIDE", R"(<w:rStyle w:val="Hid"/><w:vanish w:val="0"/>)") + "</w:p>" +
        R"(<w:p><w:pPr><w:pStyle w:val="HidPara"/></w:pPr>)" + r("HIDDEN-BY-PARA-STYLE") + "</w:p>" +
        // TOC page numbers carry webHidden: Word prints them.
        "<w:p>" + r("WEB-HIDDEN-PRINTED", "<w:webHidden/>") + "</w:p>" +
        // A field inside hidden runs: neither instruction nor result shown,
        // and text after it is back to normal.
        "<w:p>" + r("BEFORE-FIELD ") +
        R"(<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>)"
        R"(<w:r><w:rPr><w:vanish/></w:rPr><w:instrText> AUTHOR </w:instrText></w:r>)"
        R"(<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>)" +
        r("HIDDEN-FIELD-RESULT", "<w:vanish/>") +
        R"(<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>)" + r("AFTER-FIELD") +
        "</w:p>" +
        // Hidden text in a table cell.
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>)" +
        r("CELL-SHOWN ") + r("CELL-HIDDEN", "<w:vanish/>") + "</w:p></w:tc></w:tr></w:tbl>";
    const std::string text = all_text(render_pages(with_hidden_styles(body)));

    for (const char* expected : {"SHOWN-A", "SHOWN-B", "SHOWN-OVERRIDE", "WEB-HIDDEN-PRINTED",
                                 "BEFORE-FIELD", "AFTER-FIELD", "CELL-SHOWN"}) {
        INFO(expected);
        REQUIRE(contains(text, expected));
    }
    for (const char* unexpected : {"HIDDEN-DIRECT", "HIDDEN-BY-CHAR-STYLE", "HIDDEN-BY-PARA-STYLE",
                                   "HIDDEN-FIELD-RESULT", "AUTHOR", "CELL-HIDDEN"}) {
        INFO(unexpected);
        REQUIRE_FALSE(contains(text, unexpected));
    }
}

TEST_CASE("native PDF gives a fully hidden paragraph no room and no list number", "[native_pdf]") {
    const std::string numbering =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering)") + kNamespaces + ">"
        R"(<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>)"
        R"(<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>)"
        R"(<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>)";
    const std::string num_pr = R"(<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>)";
    const std::string hidden_mark = "<w:rPr><w:vanish/></w:rPr>";
    auto para = [](const std::string& ppr, const std::string& runs) {
        return "<w:p><w:pPr>" + ppr + "</w:pPr>" + runs + "</w:p>";
    };
    const std::string body =
        p("LINE-1") + p("LINE-2") +
        para(hidden_mark, r("GONE", "<w:vanish/>")) +
        para("<w:pageBreakBefore/>" + hidden_mark, r("GONE-TOO", "<w:vanish/>")) +
        p("LINE-3") +
        // Hidden content but a visible mark: an empty line stays, as in Word.
        para("", r("EMPTY-LINE-TEXT", "<w:vanish/>")) +
        p("LINE-4") +
        para(num_pr, r("ITEM-ONE")) + para(num_pr + hidden_mark, r("HIDDEN-ITEM", "<w:vanish/>")) +
        para(num_pr, r("ITEM-TWO"));
    const auto entries = render_entries(with_hidden_styles(
        body, {{"rIdN", "numbering", "numbering.xml"}}, {{"word/numbering.xml", numbering}}));

    for (const auto& e : entries) {
        INFO(e.text);
        REQUIRE(e.page == 0);  // the hidden pageBreakBefore paragraph breaks nothing
        REQUIRE_FALSE(contains(e.text, "GONE"));
        REQUIRE_FALSE(contains(e.text, "HIDDEN-ITEM"));
        REQUIRE_FALSE(contains(e.text, "3."));
    }
    const double line = entry_with(entries, "LINE-1").y - entry_with(entries, "LINE-2").y;
    REQUIRE(entry_with(entries, "LINE-2").y - entry_with(entries, "LINE-3").y == Catch::Approx(line).margin(0.5));
    REQUIRE(entry_with(entries, "LINE-3").y - entry_with(entries, "LINE-4").y ==
            Catch::Approx(2 * line).margin(0.5));
    REQUIRE(entry_with(entries, "1.").y == Catch::Approx(entry_with(entries, "ITEM-ONE").y).margin(0.5));
    REQUIRE(entry_with(entries, "2.").y == Catch::Approx(entry_with(entries, "ITEM-TWO").y).margin(0.5));
}

TEST_CASE("native PDF draws caps and small caps in capitals", "[native_pdf]") {
    const std::string styles =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
        R"(<w:style w:type="paragraph" w:styleId="CapsHead"><w:rPr><w:caps/></w:rPr></w:style>)"
        "</w:styles>";
    const std::string body =
        "<w:p>" + r("caps-latin ёжик ĳ ψ ß", "<w:caps/>") + "</w:p>" +
        R"(<w:p><w:pPr><w:pStyle w:val="CapsHead"/></w:pPr>)" + r("heading-by-style ") +
        r("not-capped", R"(<w:caps w:val="0"/>)") + "</w:p>" +
        // Right-aligned: the narrower the capitals, the further right the line starts.
        R"(<w:p><w:pPr><w:jc w:val="right"/></w:pPr>)" + r("LEAD ") + r("Small Кап", "<w:smallCaps/>") + "</w:p>" +
        R"(<w:p><w:pPr><w:jc w:val="right"/></w:pPr>)" + r("LEAD ") + r("SMALL КАП", "<w:caps/>") + "</w:p>";
    const auto entries = render_entries(make_parts(body, {{"rIdS", "styles", "styles.xml"}},
                                                   {{"word/styles.xml", styles}}, kSect));
    std::string text;
    for (const auto& e : entries) text += e.text;

    REQUIRE(contains(text, "CAPS-LATIN ЁЖИК Ĳ Ψ ß"));
    REQUIRE(contains(text, "HEADING-BY-STYLE not-capped"));
    REQUIRE_FALSE(contains(text, "Small"));
    REQUIRE_FALSE(contains(text, "Кап"));
    // Small caps are capitals, but the lower case ones are smaller: the
    // same word takes less room than in full capitals.
    REQUIRE(contains(text, "SMALL КАП"));
    std::vector<double> lead_x;  // small caps line, then caps line
    for (const auto& e : entries) {
        if (e.text.rfind("LEAD", 0) == 0) lead_x.push_back(e.x);
    }
    REQUIRE(lead_x.size() == 2);
    REQUIRE(lead_x[0] > lead_x[1] + 3.0);
}

TEST_CASE("native PDF raises superscripts and positioned text, lowers subscripts", "[native_pdf]") {
    // Default size 11 pt.
    const std::string body =
        "<w:p>" + r("BASE-SUP ") + r("SUPER", R"(<w:vertAlign w:val="superscript"/>)") + "</w:p>" +
        "<w:p>" + r("BASE-SUB ") + r("SUBSCR", R"(<w:vertAlign w:val="subscript"/>)") + "</w:p>" +
        "<w:p>" + r("BASE-POS ") + r("LOWERED", R"(<w:position w:val="-6"/>)") + "</w:p>" +
        p("LINE-A") + p("LINE-B") +
        "<w:p>" + r("LINE-C ") + r("RAISED", R"(<w:position w:val="24"/>)") + "</w:p>" +
        p("LINE-D");
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));
    auto y = [&](const char* needle) { return entry_with(entries, needle).y; };

    REQUIRE(y("SUPER") - y("BASE-SUP") == Catch::Approx(11.0 * 0.36).margin(0.1));
    REQUIRE(y("BASE-SUB") - y("SUBSCR") == Catch::Approx(11.0 * 0.12).margin(0.1));
    REQUIRE(y("BASE-POS") - y("LOWERED") == Catch::Approx(3.0).margin(0.1));
    REQUIRE(y("RAISED") - y("LINE-C") == Catch::Approx(12.0).margin(0.1));

    // Text raised 12 pt sticks out of an 11 pt line: that line grows by
    // the part that sticks out, the next one is back to normal.
    const double line = y("LINE-A") - y("LINE-B");
    REQUIRE(y("LINE-B") - y("LINE-C") > line + 5.0);
    REQUIRE(y("LINE-C") - y("LINE-D") == Catch::Approx(line).margin(0.5));
}

namespace {

struct StrokedLine {
    double      x1 = 0.0, y1 = 0.0, x2 = 0.0, y2 = 0.0;
    double      width = 1.0;
    std::string color;         // "rrggbb"
    bool        dashed = false;
};

struct DrawnImage {
    double x = 0.0, y = 0.0, w = 0.0, h = 0.0;  // y: bottom edge
    int    order = 0;                             // position among the page's drawing operations
};

struct PageDrawing {
    std::vector<std::string> fills;   // "rrggbb" of each filled rectangle, in drawing order
    std::vector<StrokedLine> lines;   // stroked straight lines, in drawing order
    std::vector<DrawnImage>  images;  // image XObjects, in drawing order
    std::vector<std::string> text_colors;  // fill color of each text show
    int first_text = -1, last_text = -1;  // order of the first/last text shown
};

// Renders `parts` and reads back what page `page`'s content stream draws.
PageDrawing render_drawing(const Parts& parts, unsigned page = 0) {
    pugi::xml_document doc;
    const std::string& xml = parts.at("word/document.xml");
    REQUIRE(doc.load_buffer(xml.data(), xml.size(), docweft::docx::kXmlParse));
    const fs::path out = temp_pdf("docweft_native_test_drawing");
    docweft::docx::render_native_pdf(doc, parts, out);

    PoDoFo::PdfMemDocument pdf;
    pdf.Load(out.string());
    PageDrawing d;
    PoDoFo::PdfContentStreamReader reader(pdf.GetPages().GetPageAt(page));
    PoDoFo::PdfContent content;
    auto hex = [](const std::vector<double>& ops) {
        char out[7];
        std::snprintf(out, sizeof out, "%02x%02x%02x", static_cast<int>(std::lround(ops[0] * 255)),
                      static_cast<int>(std::lround(ops[1] * 255)), static_cast<int>(std::lround(ops[2] * 255)));
        return std::string(out);
    };
    std::string fill = "000000";
    StrokedLine pen{0, 0, 0, 0, 1.0, "000000", false};
    int order = 0;
    std::vector<double> matrix(6, 0.0);  // last cm: an image's placement
    bool rect = false;
    while (reader.TryReadNext(content)) {
        ++order;
        if (content.GetType() == PoDoFo::PdfContentType::DoXObject) {
            d.images.push_back({matrix[4], matrix[5], matrix[0], matrix[3], order});
            continue;
        }
        if (content.GetType() != PoDoFo::PdfContentType::Operator) continue;
        std::vector<double> ops;
        for (auto it = content->Stack.rbegin(); it != content->Stack.rend(); ++it) {
            double v = 0.0;
            if (it->TryGetReal(v)) ops.push_back(v);  // names/strings (fonts, text) aren't needed
        }
        switch (content->Operator) {
            case PoDoFo::PdfOperator::rg: fill = hex(ops); break;
            case PoDoFo::PdfOperator::RG: pen.color = hex(ops); break;
            case PoDoFo::PdfOperator::w:  pen.width = ops[0]; break;
            case PoDoFo::PdfOperator::d:
                pen.dashed = false;
                for (const PoDoFo::PdfVariant& v : content->Stack) {
                    const PoDoFo::PdfArray* dashes = nullptr;
                    if (v.TryGetArray(dashes) && !dashes->IsEmpty()) pen.dashed = true;
                }
                break;
            case PoDoFo::PdfOperator::re: rect = true; break;
            case PoDoFo::PdfOperator::f:
                if (rect) d.fills.push_back(fill);
                rect = false;
                break;
            case PoDoFo::PdfOperator::m: pen.x1 = ops[0]; pen.y1 = ops[1]; break;
            case PoDoFo::PdfOperator::l: pen.x2 = ops[0]; pen.y2 = ops[1]; break;
            case PoDoFo::PdfOperator::S: d.lines.push_back(pen); break;
            case PoDoFo::PdfOperator::cm: if (ops.size() == 6) matrix = ops; break;
            case PoDoFo::PdfOperator::Tj:
            case PoDoFo::PdfOperator::TJ:
                d.text_colors.push_back(fill);
                if (d.first_text < 0) d.first_text = order;
                d.last_text = order;
                break;
            default: break;
        }
    }
    fs::remove(out);
    return d;
}

}  // namespace

TEST_CASE("native PDF applies table style conditional formatting", "[native_pdf]") {
    auto shd = [](const char* fill) {
        return std::string(R"(<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill=")") + fill + R"("/></w:tcPr>)";
    };
    const std::string styles =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
        R"(<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="400"/></w:pPr></w:pPrDefault></w:docDefaults>)"
        R"(<w:style w:type="table" w:styleId="T"><w:pPr><w:spacing w:after="0"/></w:pPr>)"
        R"(<w:tblPr><w:tblBorders><w:top w:val="nil"/><w:bottom w:val="nil"/><w:left w:val="nil"/>)"
        R"(<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr>)"
        R"(<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>)"
        R"(<w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="4"/></w:tcBorders>)"
        R"(<w:shd w:val="clear" w:color="auto" w:fill="C00000"/></w:tcPr></w:tblStylePr>)"
        R"(<w:tblStylePr w:type="lastRow">)" + shd("0000C0") + "</w:tblStylePr>"
        R"(<w:tblStylePr w:type="firstCol">)" + shd("00A000") + "</w:tblStylePr>"
        R"(<w:tblStylePr w:type="band1Horz">)" + shd("DDDDDD") + "</w:tblStylePr>"
        "</w:style></w:styles>";
    auto table = [](const std::string& look) {
        std::string rows;
        for (int r = 0; r < 3; ++r) {
            rows += "<w:tr>";
            for (int c = 0; c < 3; ++c) rows += "<w:tc>" + p("R" + std::to_string(r) + "C" + std::to_string(c)) + "</w:tc>";
            rows += "</w:tr>";
        }
        return R"(<w:tbl><w:tblPr><w:tblStyle w:val="T"/>)" + look +
               R"(</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>)" +
               rows + "</w:tbl>";
    };
    auto parts = [&](const std::string& look) {
        return make_parts(table(look), {{"rIdS", "styles", "styles.xml"}}, {{"word/styles.xml", styles}}, kSect);
    };
    auto count = [](const std::vector<std::string>& fills, const char* color) {
        return std::count(fills.begin(), fills.end(), std::string(color));
    };

    // Header row, first column, row bands on; total row off.
    const std::string look = R"(<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>)";
    const PageDrawing d = render_drawing(parts(look));
    CHECK(count(d.fills, "c00000") == 3);  // the header row, its first cell included
    CHECK(count(d.fills, "00a000") == 2);  // first column below the header
    CHECK(count(d.fills, "dddddd") == 2);  // first band = the row after the header, first column excluded
    CHECK(count(d.fills, "0000c0") == 0);  // total row not switched on
    // Only the header row's bottom border: every line at the same height.
    REQUIRE_FALSE(d.lines.empty());
    for (const StrokedLine& l : d.lines) {
        CHECK(l.y1 == Catch::Approx(d.lines.front().y1));
        CHECK(l.y2 == Catch::Approx(d.lines.front().y1));
    }

    // The table style's paragraph spacing (after = 0) beats the document default (20 pt).
    const auto entries = render_entries(parts(look));
    CHECK(entry_with(entries, "R1C1").y - entry_with(entries, "R2C1").y < 16.0);

    // Word 2007 form, bits only: noHBand | noVBand, nothing else on.
    const PageDrawing none = render_drawing(parts(R"(<w:tblLook w:val="0600"/>)"));
    CHECK(none.fills.empty());
    CHECK(none.lines.empty());
}

TEST_CASE("native PDF draws table borders with their width, color and style", "[native_pdf]") {
    // Top 2 pt red, bottom double blue, between rows 1 pt dashed green, no
    // vertical lines; cell R0C0 claims a 3 pt black bottom, which beats the
    // green top of the cell below it.
    const std::string body =
        R"(<w:tbl><w:tblPr><w:tblBorders>)"
        R"(<w:top w:val="single" w:sz="16" w:color="FF0000"/><w:bottom w:val="double" w:sz="4" w:color="0000FF"/>)"
        R"(<w:insideH w:val="dashed" w:sz="8" w:color="00FF00"/>)"
        R"(<w:left w:val="nil"/><w:right w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders></w:tblPr>)"
        R"(<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>)"
        R"(<w:tr><w:tc><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="24" w:color="auto"/></w:tcBorders></w:tcPr>)" +
        p("R0C0") + "</w:tc><w:tc>" + p("R0C1") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + p("R1C0") + "</w:tc><w:tc>" + p("R1C1") + "</w:tc></w:tr></w:tbl>";
    const PageDrawing d = render_drawing(make_parts(body, {}, {}, kSect));

    std::vector<StrokedLine> red, blue, green;
    for (const StrokedLine& l : d.lines) {
        CHECK(l.y1 == Catch::Approx(l.y2));  // horizontal only
        if (l.color == "ff0000") red.push_back(l);
        if (l.color == "0000ff") blue.push_back(l);
        if (l.color == "00ff00") green.push_back(l);
    }
    REQUIRE(red.size() == 2);
    CHECK(red[0].width == Catch::Approx(2.0));
    CHECK_FALSE(red[0].dashed);

    // Double: two thin lines per cell, one above the other.
    REQUIRE(blue.size() == 4);
    CHECK(blue[0].width == Catch::Approx(0.5));
    CHECK(std::abs(blue[0].y1 - blue[1].y1) > 1.0);

    REQUIRE_FALSE(green.empty());
    CHECK(green[0].width == Catch::Approx(1.0));
    CHECK(green[0].dashed);

    // The edge under R0C0: whatever is drawn there last is the 3 pt black line.
    const double col0_x = red[0].x1;
    const double edge_y = green[0].y1;
    const StrokedLine* last = nullptr;
    for (const StrokedLine& l : d.lines) {
        if (l.x1 == Catch::Approx(col0_x) && l.y1 == Catch::Approx(edge_y)) last = &l;
    }
    REQUIRE(last != nullptr);
    CHECK(last->color == "000000");
    CHECK(last->width == Catch::Approx(3.0));
    CHECK_FALSE(last->dashed);
}

namespace {

std::vector<StrokedLine> lines_in(const PageDrawing& d, const std::string& color, bool horizontal) {
    std::vector<StrokedLine> out;
    for (const StrokedLine& l : d.lines) {
        if (l.color == color && (std::abs(l.y1 - l.y2) < 0.01) == horizontal) out.push_back(l);
    }
    return out;
}

// All four sides, `sz` eighths of a point, `space` points from the text.
std::string p_bdr(const char* color, int sz, int space, const std::string& between = {}) {
    std::string sides;
    for (const char* side : {"top", "left", "bottom", "right"}) {
        sides += std::string("<w:") + side + R"( w:val="single" w:sz=")" + std::to_string(sz) + R"(" w:space=")" +
                 std::to_string(space) + R"(" w:color=")" + color + R"("/>)";
    }
    return "<w:pBdr>" + sides + between + "</w:pBdr>";
}

}  // namespace

TEST_CASE("native PDF draws paragraph borders and shading, grouping equal ones", "[native_pdf]") {
    const std::string box = p_bdr("0000FF", 8, 4) + R"(<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>)";
    const std::string group =
        p_bdr("FF0000", 8, 2, R"(<w:between w:val="single" w:sz="4" w:space="1" w:color="00FF00"/>)");
    const std::string body = p("BEFORE") + p("BOXED", box) + p("MIDDLE") + p("G-ONE", group) +
                             p("G-TWO", group) + p("G-THREE", group) + p("AFTER");
    const Parts parts = make_parts(body, {}, {}, kSect);
    const PageDrawing d = render_drawing(parts);
    const auto entries = render_entries(parts);
    auto y = [&](const char* needle) { return entry_with(entries, needle).y; };

    // Single box: shaded, one line per side; left/right outside the text
    // area by space + width (the border line centred in its width).
    CHECK(std::count(d.fills.begin(), d.fills.end(), std::string("fff2cc")) > 0);
    const auto blue_h = lines_in(d, "0000ff", true), blue_v = lines_in(d, "0000ff", false);
    REQUIRE(blue_h.size() == 2);
    REQUIRE(blue_v.size() == 2);
    CHECK(std::max(blue_h[0].y1, blue_h[1].y1) > y("BOXED"));
    CHECK(std::min(blue_h[0].y1, blue_h[1].y1) < y("BOXED"));
    CHECK(std::min(blue_v[0].x1, blue_v[1].x1) == Catch::Approx(56.7 - 4.0 - 1.0 + 0.5).margin(0.1));
    CHECK(std::max(blue_v[0].x1, blue_v[1].x1) == Catch::Approx(595.3 - 56.7 + 4.0 + 1.0 - 0.5).margin(0.1));
    // The box takes room: space + width above and below the text.
    const double line = y("MIDDLE") - y("G-ONE") - 3.0;  // G-ONE's own top pad: 2 + 1
    CHECK(y("BEFORE") - y("BOXED") == Catch::Approx(line + 5.0).margin(0.5));
    CHECK(y("BOXED") - y("MIDDLE") == Catch::Approx(line + 5.0).margin(0.5));

    // Group: top and bottom once, "between" lines between its paragraphs,
    // sides unbroken from the first paragraph to the last.
    CHECK(lines_in(d, "ff0000", true).size() == 2);
    CHECK(lines_in(d, "00ff00", true).size() == 2);
    std::vector<StrokedLine> left;
    for (const StrokedLine& l : lines_in(d, "ff0000", false)) {
        if (l.x1 < 100.0) left.push_back(l);
    }
    REQUIRE(left.size() == 3);
    std::sort(left.begin(), left.end(), [](const StrokedLine& a, const StrokedLine& b) { return a.y2 > b.y2; });
    CHECK(left[0].y1 == Catch::Approx(left[1].y2).margin(0.01));
    CHECK(left[1].y1 == Catch::Approx(left[2].y2).margin(0.01));
}

TEST_CASE("native PDF continues a paragraph box across a page break", "[native_pdf]") {
    std::string text;
    for (int i = 0; i < 1400; ++i) text += "word ";
    const Parts parts = make_parts(p(text, p_bdr("0000FF", 8, 4)), {}, {}, kSect);
    const PageDrawing first = render_drawing(parts, 0);
    const PageDrawing second = render_drawing(parts, 1);

    // Top border on the first page, bottom on the second, sides on both.
    const auto first_h = lines_in(first, "0000ff", true), second_h = lines_in(second, "0000ff", true);
    REQUIRE(first_h.size() == 1);
    REQUIRE(second_h.size() == 1);
    CHECK(first_h[0].y1 > 700.0);
    CHECK(second_h[0].y1 < first_h[0].y1);
    CHECK(lines_in(first, "0000ff", false).size() == 2);
    CHECK(lines_in(second, "0000ff", false).size() == 2);
}

TEST_CASE("native PDF resolves theme colors with tint and shade like Word", "[native_pdf]") {
    // No theme part: Office's default theme, accent1 = 4472C4. Expected
    // values are Word's own palette entries for it.
    auto shaded = [](const char* text, const std::string& shd) {
        return p(text, R"(<w:shd w:val="clear" w:color="auto" )" + shd + "/>");
    };
    auto boxed = [](const char* text, const std::string& color_attrs) {
        return p(text, R"(<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" )" + color_attrs + "/></w:pBdr>");
    };
    const std::string body =
        shaded("DARKER-25", R"(w:themeFill="accent1" w:themeFillShade="BF")") + p("GAP") +
        shaded("LIGHTER-80", R"(w:themeFill="accent1" w:themeFillTint="33")") + p("GAP") +
        shaded("HEX-WINS", R"(w:fill="FF0000" w:themeFill="accent1")") + p("GAP") +
        boxed("GRID-LINE", R"(w:themeColor="accent1" w:themeTint="99")") +
        boxed("TEXT2", R"(w:themeColor="text2")") + boxed("DARK1", R"(w:themeColor="dark1")");
    const PageDrawing d = render_drawing(make_parts(body, {}, {}, kSect));

    auto has_fill = [&](const char* c) { return std::count(d.fills.begin(), d.fills.end(), std::string(c)) > 0; };
    CHECK(has_fill("2f5496"));
    CHECK(has_fill("d9e2f3"));
    CHECK(has_fill("ff0000"));
    CHECK_FALSE(has_fill("4472c4"));
    CHECK(lines_in(d, "8eaadb", true).size() == 1);
    CHECK(lines_in(d, "44546a", true).size() == 1);
    CHECK(lines_in(d, "000000", true).size() == 1);
}

TEST_CASE("native PDF fills in TOC page numbers, PAGEREF and NUMPAGES from its own layout", "[native_pdf]") {
    auto field = [](const std::string& instr, const std::string& cached) {
        return R"(<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> )" + instr +
               R"( </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>)" + cached +
               R"(</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>)";
    };
    // Word's TOC: one field around the entries, each entry a hyperlink
    // with a PAGEREF to its heading's bookmark; cached numbers are stale.
    auto entry = [&](const std::string& text, const std::string& bookmark, const std::string& cached) {
        return R"(<w:hyperlink w:anchor=")" + bookmark + R"("><w:r><w:t>)" + text +
               R"(</w:t></w:r><w:r><w:tab/></w:r>)" + field("PAGEREF " + bookmark + R"( \h)", cached) + "</w:hyperlink>";
    };
    const std::string toc_tabs = R"(<w:pPr><w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs></w:pPr>)";
    auto heading = [](const std::string& text, const std::string& bookmark, int id) {
        return R"(<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:bookmarkStart w:id=")" + std::to_string(id) +
               R"(" w:name=")" + bookmark + R"("/><w:r><w:t>)" + text + R"(</w:t></w:r><w:bookmarkEnd w:id=")" +
               std::to_string(id) + R"("/></w:p>)";
    };
    const std::string body =
        "<w:p>" + toc_tabs + R"(<w:r><w:fldChar w:fldCharType="begin"/></w:r>)"
        R"(<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>)"
        R"(<w:r><w:fldChar w:fldCharType="separate"/></w:r>)" + entry("Entry-one", "_Toc1", "99") + "</w:p>" +
        "<w:p>" + toc_tabs + entry("Entry-two", "_Toc2", "98") + "</w:p>" +
        "<w:p>" + toc_tabs + entry("Entry-cell", "_Toc3", "97") +
        R"(<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>)" +
        "<w:p><w:r><w:t xml:space=\"preserve\">Body page </w:t></w:r>" + field("PAGE", "8") +
        "<w:r><w:t xml:space=\"preserve\"> of </w:t></w:r>" + field("NUMPAGES", "9") + "</w:p>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">Missing ref </w:t></w:r>" + field("PAGEREF _Nowhere \\h", "77") + "</w:p>" +
        heading("Chapter one", "_Toc1", 1) +
        R"(<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Table page</w:t></w:r></w:p>)"
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:bookmarkStart w:id="3" w:name="_Toc3"/>)"
        R"(<w:r><w:t>In a cell</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p></w:tc></w:tr></w:tbl>)" +
        heading("Chapter two", "_Toc2", 2);
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));

    // Everything on the line of an entry, in reading order.
    auto line_of = [&](const std::string& needle) {
        const TextEntry& e = entry_with(entries, needle);
        std::string text;
        for (const auto& other : entries) {
            if (other.page == e.page && std::abs(other.y - e.y) < 0.5) text += other.text;
        }
        return text;
    };
    CHECK(entry_with(entries, "Chapter one").page == 1);
    CHECK(entry_with(entries, "In a cell").page == 2);
    CHECK(entry_with(entries, "Chapter two").page == 3);
    CHECK(line_of("Entry-one").back() == '2');
    CHECK(line_of("Entry-cell").back() == '3');
    CHECK(line_of("Entry-two").back() == '4');
    CHECK(contains(line_of("Body page"), "Body page 1 of 4"));
    CHECK(contains(line_of("Missing ref"), "77"));
    for (const auto& e : entries) {
        INFO(e.text);
        CHECK_FALSE(contains(e.text, "99"));
        CHECK_FALSE(contains(e.text, "98"));
        CHECK_FALSE(contains(e.text, "97"));
    }
}

TEST_CASE("native PDF draws tab leaders", "[native_pdf]") {
    const std::string styles =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
        R"(<w:style w:type="paragraph" w:styleId="Toc"><w:pPr><w:tabs>)"
        R"(<w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr></w:style></w:styles>)";
    auto line = [](const std::string& ppr, const std::string& left, const std::string& right) {
        return "<w:p><w:pPr>" + ppr + "</w:pPr><w:r><w:t>" + left + "</w:t></w:r><w:r><w:tab/><w:t>" + right +
               "</w:t></w:r></w:p>";
    };
    const std::string toc = R"(<w:pStyle w:val="Toc"/>)";
    const std::string body =
        line(toc, "Dotted", "11") +
        // Same position redefined by the paragraph: its leader replaces the style's.
        line(toc + R"(<w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="9000"/></w:tabs>)", "Hyphened", "12") +
        line(R"(<w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs>)", "Plain", "13") +
        // A bar tab isn't a stop: the text still goes to the dotted right tab.
        line(R"(<w:tabs><w:tab w:val="bar" w:pos="2000"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs>)",
             "Barred", "14");
    const auto entries = render_entries(make_parts(body, {{"rIdS", "styles", "styles.xml"}},
                                                   {{"word/styles.xml", styles}}, kSect));
    auto line_text = [&](const std::string& needle) {
        const TextEntry& e = entry_with(entries, needle);
        std::string text;
        for (const auto& other : entries) {
            if (other.page == e.page && std::abs(other.y - e.y) < 0.5) text += other.text;
        }
        return text;
    };
    auto count = [](const std::string& text, char c) { return std::count(text.begin(), text.end(), c); };

    const std::string dotted = line_text("Dotted");
    CHECK(count(dotted, '.') > 40);  // most of a 15 cm gap
    CHECK(contains(dotted, "11"));
    const std::string hyphened = line_text("Hyphened");
    CHECK(count(hyphened, '-') > 20);
    CHECK(count(hyphened, '.') == 0);
    CHECK(count(line_text("Plain"), '.') == 0);
    CHECK(count(line_text("Barred"), '.') > 40);
}

namespace {

// A floating picture (<wp:anchor>) of `w` x `h` points; `hpos`/`vpos`:
// "<wp:posOffset>…</wp:posOffset>" or "<wp:align>…</wp:align>".
std::string floating(double w, double h, const std::string& hrel, const std::string& hpos,
                     const std::string& vrel, const std::string& vpos, const std::string& wrap,
                     bool behind = false) {
    const std::string cx = std::to_string(static_cast<long long>(w * 12700));
    const std::string cy = std::to_string(static_cast<long long>(h * 12700));
    return R"(<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc=")" +
           std::string(behind ? "1" : "0") + R"(" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>)"
           R"(<wp:positionH relativeFrom=")" + hrel + R"(">)" + hpos + "</wp:positionH>"
           R"(<wp:positionV relativeFrom=")" + vrel + R"(">)" + vpos + "</wp:positionV>"
           R"(<wp:extent cx=")" + cx + R"(" cy=")" + cy + R"("/>)" + wrap +
           R"(<wp:docPr id="1" name="f"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">)"
           R"(<pic:pic><pic:blipFill><a:blip r:embed="rIdP"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>)"
           "</wp:anchor></w:drawing></w:r>";
}

std::string offset(double pt) {
    return "<wp:posOffset>" + std::to_string(static_cast<long long>(pt * 12700)) + "</wp:posOffset>";
}

Parts with_picture(const std::string& body, const std::vector<Rel>& rels = {}, Parts extra = {},
                   const std::string& sect = kSect) {
    std::vector<Rel> all = rels;
    all.push_back({"rIdP", "image", "media/image1.png"});
    extra["word/media/image1.png"] = std::string(tf_test::tiny_png_bytes());
    return make_parts(body, all, extra, sect);
}

std::string para_with(const std::string& text, const std::string& runs = {}) {
    return "<w:p>" + runs + R"(<w:r><w:t xml:space="preserve">)" + text + "</w:t></w:r></w:p>";
}

constexpr double kPageHeight = 841.9;  // A4, kSect

}  // namespace

TEST_CASE("native PDF places floating pictures in front of and behind the text", "[native_pdf]") {
    const std::string body =
        para_with("TEXT-A", floating(50, 40, "page", offset(100), "page", offset(200), "<wp:wrapNone/>")) +
        para_with("TEXT-B", floating(60, 30, "margin", "<wp:align>right</wp:align>", "paragraph", offset(10),
                                     "<wp:wrapNone/>", /*behind*/ true)) +
        para_with("TEXT-C");
    const Parts parts = with_picture(body);
    const PageDrawing d = render_drawing(parts);
    const auto entries = render_entries(parts);
    auto y = [&](const char* needle) { return entry_with(entries, needle).y; };

    REQUIRE(d.images.size() == 2);
    const DrawnImage& front = d.images[0].w == Catch::Approx(50.0) ? d.images[0] : d.images[1];
    const DrawnImage& behind = &front == &d.images[0] ? d.images[1] : d.images[0];
    // In front: page-relative, 100 pt from the left, 200 pt from the top.
    CHECK(front.x == Catch::Approx(100.0).margin(0.1));
    CHECK(front.y == Catch::Approx(kPageHeight - 200.0 - 40.0).margin(0.2));
    CHECK(front.order > d.last_text);
    // Behind: right-aligned in the margins (right margin at 595.3 - 56.7),
    // drawn before any text of the page.
    CHECK(behind.x + behind.w == Catch::Approx(595.3 - 56.7).margin(0.2));
    CHECK(behind.order < d.first_text);
    // wrapNone takes no room: the lines follow each other as usual.
    CHECK(y("TEXT-A") - y("TEXT-B") == Catch::Approx(y("TEXT-B") - y("TEXT-C")).margin(0.5));
}

TEST_CASE("native PDF keeps text out of a top-and-bottom floating picture's band", "[native_pdf]") {
    std::string body = para_with("BEFORE") +
        para_with("ANCHOR", floating(80, 100, "margin", "<wp:align>center</wp:align>", "paragraph", offset(0),
                                     "<wp:wrapTopAndBottom/>"));
    // A second picture, positioned on the page 300..380 pt from the top,
    // anchored in the first line: the lines below must skip that band.
    body = para_with("FIRST", floating(100, 80, "page", offset(200), "page", offset(300), "<wp:wrapTopAndBottom/>")) + body;
    for (int i = 0; i < 25; ++i) body += para_with("LINE-" + std::to_string(i));
    const auto entries = render_entries(with_picture(body));
    auto y = [&](const char* needle) { return entry_with(entries, needle).y; };

    const double line = y("LINE-0") - y("LINE-1");
    // The paragraph-relative picture sits between BEFORE and its paragraph's text.
    CHECK(y("BEFORE") - y("ANCHOR") == Catch::Approx(line + 100.0).margin(0.5));
    const double band_top = kPageHeight - 300.0, band_bottom = kPageHeight - 380.0;
    bool below = false;
    for (const auto& e : entries) {
        INFO(e.text << " at " << e.y);
        CHECK_FALSE((e.y > band_bottom && e.y < band_top));  // no baseline inside the band
        below = below || e.y < band_bottom;
    }
    CHECK(below);
}

TEST_CASE("native PDF flows text beside square-wrapped floating pictures", "[native_pdf]") {
    constexpr double kLeft = 56.7, kRight = 595.3 - 56.7;  // kSect's margins
    std::string words;
    for (int i = 0; i < 160; ++i) words += "word" + std::to_string(i) + " ";

    SECTION("beside it while the lines meet it, the whole width below") {
        // 150 x 120 pt at the left margin, the paragraph's top: text on its right.
        const std::string body =
            para_with(words, floating(150, 120, "margin", "<wp:align>left</wp:align>", "paragraph", offset(0),
                                      R"(<wp:wrapSquare wrapText="bothSides"/>)")) +
            para_with("NEXT-PARA");
        const auto entries = render_entries(with_picture(body));
        const double top = entry_with(entries, "word0").y + 12.0, bottom = top - 120.0;
        int beside = 0, below = 0;
        for (const auto& e : entries) {
            INFO(e.text << " at " << e.x << ", " << e.y);
            if (e.page != 0) continue;
            if (e.y > bottom) {
                CHECK(e.x >= kLeft + 150.0 - 0.5);
                ++beside;
            } else if (e.x < kLeft + 1.0) {
                ++below;
            }
        }
        CHECK(beside >= 5);
        CHECK(below >= 3);
        CHECK(entry_with(entries, "NEXT-PARA").x == Catch::Approx(kLeft).margin(0.5));
    }
    SECTION("wrapText=left keeps the text on the left") {
        const std::string body =
            para_with(words, floating(150, 120, "page", offset(250), "paragraph", offset(0),
                                      R"(<wp:wrapSquare wrapText="left"/>)"));
        const auto entries = render_entries(with_picture(body));
        const double bottom = entry_with(entries, "word0").y + 12.0 - 120.0;
        for (const auto& e : entries) {
            INFO(e.text << " at " << e.x << ", " << e.y);
            if (e.page == 0 && e.y > bottom) CHECK(e.x < 250.0);
        }
    }
    SECTION("no room beside it: the lines move below") {
        const std::string body =
            para_with("FIRST") +
            para_with(words, floating(kRight - kLeft - 10.0, 100, "margin", "<wp:align>left</wp:align>",
                                      "paragraph", offset(0), R"(<wp:wrapTight wrapText="bothSides"/>)"));
        const auto entries = render_entries(with_picture(body));
        const double top = entry_with(entries, "FIRST").y - 4.0;
        const double bottom = top - 100.0;
        for (const auto& e : entries) {
            INFO(e.text << " at " << e.y);
            if (e.page == 0 && e.text.find("FIRST") == std::string::npos) CHECK(e.y < bottom);
        }
        CHECK(entry_with(entries, "word0").y == Catch::Approx(bottom - 10.0).margin(5.0));
    }
    SECTION("a table after it goes below it") {
        const std::string body =
            para_with("ANCHOR", floating(150, 200, "margin", "<wp:align>left</wp:align>", "paragraph", offset(0),
                                         R"(<wp:wrapSquare wrapText="bothSides"/>)")) +
            R"(<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>CELL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>)";
        const auto entries = render_entries(with_picture(body));
        CHECK(entry_with(entries, "CELL").y < entry_with(entries, "ANCHOR").y + 12.0 - 200.0);
    }
}

TEST_CASE("native PDF doesn't let a floating header logo push the body down", "[native_pdf]") {
    auto render_first_line = [&](const std::string& header_runs) {
        const std::string header =
            std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr)") + kNamespaces + ">" +
            para_with("HEADER", header_runs) + "</w:hdr>";
        const std::string sect =
            R"(<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/>)"
            R"(<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567"/></w:sectPr>)";
        const std::string header_rels =
            R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
            R"(<Relationship Id="rIdP" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>)";
        const Parts parts = with_picture(para_with("BODY"), {{"rIdH", "header", "header1.xml"}},
                                         {{"word/header1.xml", header}, {"word/_rels/header1.xml.rels", header_rels}}, sect);
        return entry_with(render_entries(parts), "BODY").y;
    };
    const double plain = render_first_line("");
    const double with_logo =
        render_first_line(floating(80, 120, "page", offset(20), "page", offset(10), "<wp:wrapNone/>"));
    CHECK(with_logo == Catch::Approx(plain).margin(0.1));
}

TEST_CASE("native PDF formats page numbers per section and per field switch", "[native_pdf]") {
    auto field = [](const std::string& instr, const std::string& cached) {
        return R"(<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> )" + instr +
               R"( </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>)" + cached +
               R"(</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>)";
    };
    auto labelled = [&](const std::string& label, const std::string& instr) {
        return R"(<w:p><w:r><w:t xml:space="preserve">)" + label + " </w:t></w:r>" + field(instr, "0") + "</w:p>";
    };
    auto bookmarked = [](const std::string& text, const std::string& name) {
        return R"(<w:p><w:bookmarkStart w:id="1" w:name=")" + name + R"("/><w:r><w:t>)" + text +
               R"(</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>)";
    };
    const std::string footer =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr)") + kNamespaces + ">" +
        labelled("Folio", "PAGE") + "</w:ftr>";
    const std::string page_break = R"(<w:p><w:r><w:br w:type="page"/></w:r></w:p>)";
    // Front matter: lower-case roman, two pages; then the body: decimal from 1.
    const std::string front_sect =
        R"(<w:p><w:pPr><w:sectPr><w:footerReference w:type="default" r:id="rIdF"/><w:pgNumType w:fmt="lowerRoman"/>)"
        R"(<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:footer="567"/>)"
        "</w:sectPr></w:pPr></w:p>";
    const std::string body =
        labelled("Ref-front", "PAGEREF _Front \\h") + labelled("Ref-body", "PAGEREF _Body \\h") +
        labelled("Total", "NUMPAGES \\* roman") + labelled("Dash", "PAGE \\* ArabicDash \\* MERGEFORMAT") +
        page_break + bookmarked("Front-two", "_Front") + front_sect +
        bookmarked("Body-one", "_Body") + labelled("Upper", "PAGE \\* ROMAN");
    const std::string last_sect =
        R"(<w:sectPr><w:pgNumType w:start="1"/><w:pgSz w:w="11906" w:h="16838"/>)"
        R"(<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:footer="567"/></w:sectPr>)";
    const auto entries = render_entries(
        make_parts(body, {{"rIdF", "footer", "footer1.xml"}}, {{"word/footer1.xml", footer}}, last_sect));

    auto line_of = [&](const std::string& needle, int page = -1) {
        for (const auto& e : entries) {
            if (!contains(e.text, needle) || (page >= 0 && e.page != page)) continue;
            std::string text;
            for (const auto& other : entries) {
                if (other.page == e.page && std::abs(other.y - e.y) < 0.5) text += other.text;
            }
            return text;
        }
        FAIL("no line with " << needle);
        return std::string();
    };
    CHECK(contains(line_of("Folio", 0), "Folio i"));
    CHECK(contains(line_of("Folio", 1), "Folio ii"));
    CHECK(contains(line_of("Folio", 2), "Folio 1"));
    CHECK(contains(line_of("Ref-front"), "Ref-front ii"));  // in the bookmark's section's format
    CHECK(contains(line_of("Ref-body"), "Ref-body 1"));
    CHECK(contains(line_of("Total"), "Total iii"));         // NUMPAGES with its own switch
    CHECK(contains(line_of("Dash"), "Dash - 1 -"));
    CHECK(contains(line_of("Upper"), "Upper I"));
}

TEST_CASE("native PDF draws chart data labels as the chart asks", "[native_pdf]") {
    // One series over Q1..Q4 in a <c:{type}> group, with `labels` (a
    // <c:dLbls>) inside the series.
    auto chart = [](const std::string& type, const std::string& head, const std::vector<std::string>& values,
                    const std::string& labels, const std::string& format = "General") {
        std::string cats, vals;
        for (std::size_t i = 0; i < values.size(); ++i) {
            cats += R"(<c:pt idx=")" + std::to_string(i) + R"("><c:v>Q)" + std::to_string(i + 1) + "</c:v></c:pt>";
            vals += R"(<c:pt idx=")" + std::to_string(i) + R"("><c:v>)" + values[i] + "</c:v></c:pt>";
        }
        const bool pie = type == "pieChart";
        return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") + kNamespaces +
               R"(><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:)" + type + ">" + head +
               R"(<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series</c:v></c:pt></c:strCache></c:strRef></c:tx>)" +
               labels + "<c:cat><c:strRef><c:strCache>" + cats + "</c:strCache></c:strRef></c:cat>" +
               "<c:val><c:numRef><c:numCache><c:formatCode>" + format + "</c:formatCode>" + vals +
               "</c:numCache></c:numRef></c:val></c:ser>" +
               (pie ? "</c:pieChart>" : R"(<c:axId val="1"/><c:axId val="2"/></c:)" + type +
                                        R"(><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>)") +
               "</c:plotArea></c:chart></c:chartSpace>";
    };
    auto render_text = [&](const std::string& chart_part) {
        const Parts parts = make_parts(chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
                                       {{"word/charts/chart1.xml", chart_part}}, kSect);
        return all_text(render_pages(parts));
    };
    const std::string bar_head = R"(<c:barDir val="col"/><c:grouping val="clustered"/>)";

    // Values in the labels' own number format; one point's label deleted.
    const std::string bar = render_text(chart("barChart", bar_head, {"1200", "2350", "1800", "3100"},
        R"(<c:dLbls><c:dLbl><c:idx val="2"/><c:delete val="1"/></c:dLbl><c:numFmt formatCode="#,##0" sourceLinked="0"/>)"
        R"(<c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>)"));
    CHECK(contains(bar, "1,200"));
    CHECK(contains(bar, "3,100"));
    CHECK_FALSE(contains(bar, "1,800"));

    // No <c:dLbls>: no labels at all (the pie used to get percentages anyway).
    const std::string plain_pie = render_text(chart("pieChart", "", {"45", "30", "15", "10"}, ""));
    CHECK_FALSE(contains(plain_pie, "%"));

    // Category and percentage, Office's separator; a point with its own text.
    const std::string pie = render_text(chart("pieChart", "", {"45", "30", "15", "10"},
        R"(<c:dLbls><c:dLbl><c:idx val="3"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Other</a:t></a:r></a:p></c:rich></c:tx>)"
        R"(<c:showVal val="1"/></c:dLbl><c:dLblPos val="outEnd"/><c:showCatName val="1"/><c:showPercent val="1"/></c:dLbls>)"));
    CHECK(contains(pie, "Q1, 45%"));
    CHECK(contains(pie, "Q3, 15%"));
    CHECK(contains(pie, "Other"));
    CHECK_FALSE(contains(pie, "Q4, 10%"));

    // The values' own (source-linked) format; halves round away from zero, as in Excel.
    const std::string line = render_text(chart("lineChart", R"(<c:grouping val="standard"/>)", {"3.5", "4.25", "2.75", "5"},
        R"(<c:dLbls><c:dLblPos val="t"/><c:showVal val="1"/></c:dLbls>)", "0.0"));
    CHECK(contains(line, "4.3"));
    CHECK(contains(line, "2.8"));
    CHECK(contains(line, "5.0"));
}

TEST_CASE("native PDF leaves out pictures in hidden runs", "[native_pdf]") {
    auto picture = [](const std::string& rpr) {
        return R"(<w:r><w:rPr>)" + rpr + R"(</w:rPr><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>)"
               R"(<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill>)"
               R"(<a:blip r:embed="rIdP"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>)";
    };
    auto image_count = [&](const std::string& rpr) {
        const Parts parts = with_hidden_styles("<w:p>" + r("TEXT ") + picture(rpr) + "</w:p>",
                                               {{"rIdP", "image", "media/image1.png"}},
                                               {{"word/media/image1.png", std::string(tf_test::tiny_png_bytes())}});
        pugi::xml_document doc;
        const std::string& xml = parts.at("word/document.xml");
        REQUIRE(doc.load_buffer(xml.data(), xml.size(), docweft::docx::kXmlParse));
        const fs::path out = temp_pdf("docweft_native_test_hidden_image");
        docweft::docx::render_native_pdf(doc, parts, out);
        PoDoFo::PdfMemDocument pdf;
        pdf.Load(out.string());
        int n = 0;
        for (const PoDoFo::PdfObject* obj : pdf.GetObjects()) {
            if (!obj->IsDictionary()) continue;
            const PoDoFo::PdfObject* subtype = obj->GetDictionary().GetKey("Subtype");
            if (subtype != nullptr && subtype->IsName() && subtype->GetName() == "Image") ++n;
        }
        fs::remove(out);
        return n;
    };
    REQUIRE(image_count("") == 1);
    REQUIRE(image_count("<w:vanish/>") == 0);
}

TEST_CASE("native PDF writes chart numbers with the document's separators", "[native_pdf]") {
    // Bar values labelled "#,##0.0"; the value axis steps by 0.2 for the
    // second chart. `lang`: the chart's <c:lang>, `doc_lang`: styles.xml's.
    auto render_text = [](double q1, double q2, const std::string& lang, const std::string& doc_lang) {
        auto num = [](double v) { std::ostringstream os; os << v; return os.str(); };
        const std::string chart =
            std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") + kNamespaces + ">" +
            (lang.empty() ? "" : R"(<c:lang val=")" + lang + R"("/>)") +
            R"(<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)"
            R"(<c:ser><c:idx val="0"/><c:order val="0"/>)"
            R"(<c:dLbls><c:numFmt formatCode="#,##0.0" sourceLinked="0"/><c:showVal val="1"/></c:dLbls>)"
            R"(<c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>)"
            R"(<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>)" + num(q1) + R"(</c:v></c:pt><c:pt idx="1"><c:v>)" + num(q2) +
            R"(</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes +
            "</c:plotArea></c:chart></c:chartSpace>";
        Parts extra{{"word/charts/chart1.xml", chart}};
        if (!doc_lang.empty()) {
            extra["word/styles.xml"] =
                std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">" +
                R"(<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val=")" + doc_lang +
                R"("/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>)";
        }
        return all_text(render_pages(make_parts(chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}}, extra, kSect)));
    };
    const std::string nbsp = " ";

    const std::string ru = render_text(1234.5, 2.25, "ru-RU", "");
    CHECK(contains(ru, "1" + nbsp + "234,5"));
    CHECK(contains(ru, "2,3"));  // 2.25 rounded half up, decimal comma
    CHECK_FALSE(contains(ru, "1,234.5"));
    CHECK(contains(render_text(0.3, 0.7, "ru-RU", ""), "0,2"));  // axis labels too

    // No <c:lang>: the document's language decides.
    CHECK(contains(render_text(1234.5, 2.25, "", "de-DE"), "1.234,5"));
    CHECK(contains(render_text(1234.5, 2.25, "", "ru-RU"), "1" + nbsp + "234,5"));
    // The chart's own language wins; English (and no language at all) as before.
    CHECK(contains(render_text(1234.5, 2.25, "en-US", "ru-RU"), "1,234.5"));
    const std::string none = render_text(0.3, 0.7, "", "");
    CHECK(contains(none, "0.2"));
}

TEST_CASE("native PDF follows a table style's basedOn chain", "[native_pdf]") {
    // Base: red inside lines, green header row, 1 cm left cell margin.
    // Derived (used by the table): blue top border, a blue first column,
    // and its own header rows bold on top of the base's green.
    const std::string styles =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
        R"(<w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblBorders>)"
        R"(<w:top w:val="nil"/><w:bottom w:val="nil"/><w:left w:val="nil"/><w:right w:val="nil"/>)"
        R"(<w:insideH w:val="single" w:sz="8" w:color="FF0000"/><w:insideV w:val="nil"/></w:tblBorders>)"
        R"(<w:tblCellMar><w:left w:w="567" w:type="dxa"/></w:tblCellMar></w:tblPr>)"
        R"(<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="00A000"/></w:tcPr></w:tblStylePr>)"
        R"(</w:style>)"
        R"(<w:style w:type="table" w:styleId="Derived"><w:basedOn w:val="Base"/><w:tblPr><w:tblBorders>)"
        R"(<w:top w:val="single" w:sz="8" w:color="0000FF"/></w:tblBorders></w:tblPr>)"
        R"(<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>)"
        R"(<w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="0000C0"/></w:tcPr></w:tblStylePr>)"
        R"(</w:style></w:styles>)";
    std::string rows;
    for (int r = 0; r < 3; ++r) {
        rows += "<w:tr>";
        for (int c = 0; c < 2; ++c) rows += "<w:tc>" + p("R" + std::to_string(r) + "C" + std::to_string(c)) + "</w:tc>";
        rows += "</w:tr>";
    }
    const std::string body =
        R"(<w:tbl><w:tblPr><w:tblStyle w:val="Derived"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:noHBand="1" w:noVBand="1"/></w:tblPr>)"
        R"(<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>)" + rows + "</w:tbl>";
    const Parts parts = make_parts(body, {}, {{"word/styles.xml", styles}}, kSect);
    const PageDrawing d = render_drawing(parts);

    CHECK(std::count(d.fills.begin(), d.fills.end(), std::string("00a000")) == 2);  // header row: the base's region
    CHECK(std::count(d.fills.begin(), d.fills.end(), std::string("0000c0")) == 2);  // first column below it
    CHECK(lines_in(d, "0000ff", true).size() == 2);   // the derived style's top border
    std::set<long> red_rows;  // the base's lines between the rows (a shared edge is stroked twice)
    for (const StrokedLine& l : lines_in(d, "ff0000", true)) red_rows.insert(std::lround(l.y1));
    CHECK(red_rows.size() == 2);
    CHECK(lines_in(d, "000000", false).empty());      // base says: no vertical lines
    // The base's 1 cm left cell margin.
    CHECK(entry_with(render_entries(parts), "R1C0").x == Catch::Approx(56.7 + 28.35).margin(1.0));
}

TEST_CASE("native PDF lets a table style's font size and alignment beat Normal's 11/12 pt left", "[native_pdf]") {
    // Normal: 12 pt, left. Table style: 8 pt, centred. "Body": 12 pt, left,
    // but not the default style — it always wins over the table style.
    const std::string styles =
        std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles)") + kNamespaces + ">"
        R"(<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:pPr><w:jc w:val="left"/></w:pPr>)"
        R"(<w:rPr><w:sz w:val="24"/></w:rPr></w:style>)"
        R"(<w:style w:type="paragraph" w:styleId="Body"><w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>)"
        R"(<w:style w:type="table" w:styleId="T"><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="16"/></w:rPr></w:style>)"
        "</w:styles>";
    const std::string cell_paras = p("NORMAL-A") + p("NORMAL-B") + p("BODY-A", R"(<w:pStyle w:val="Body"/>)") +
                                   p("BODY-B", R"(<w:pStyle w:val="Body"/>)");
    const std::string body =
        R"(<w:tbl><w:tblPr><w:tblStyle w:val="T"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc>)" +
        cell_paras + "</w:tc></w:tr></w:tbl>";
    auto render = [&](const std::string& compat) {
        Parts extra{{"word/styles.xml", styles}};
        if (!compat.empty()) {
            extra["word/settings.xml"] =
                std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings)") + kNamespaces +
                "><w:compat>" + compat + "</w:compat></w:settings>";
        }
        return render_entries(make_parts(body, {}, extra, kSect));
    };
    auto gap = [](const std::vector<TextEntry>& e, const char* a, const char* b) {
        return entry_with(e, a).y - entry_with(e, b).y;
    };

    // No compat setting (older documents): the table style wins against Normal.
    const auto quirk = render("");
    CHECK(gap(quirk, "NORMAL-A", "NORMAL-B") < gap(quirk, "BODY-A", "BODY-B") - 3.0);  // 8 pt lines vs 12 pt
    CHECK(entry_with(quirk, "NORMAL-A").x > entry_with(quirk, "BODY-A").x + 50.0);    // centred vs left

    // Word 2013+ documents opt out: Normal wins like any paragraph style.
    const auto modern = render(
        R"(<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>)");
    CHECK(gap(modern, "NORMAL-A", "NORMAL-B") == Catch::Approx(gap(modern, "BODY-A", "BODY-B")).margin(0.1));
    CHECK(entry_with(modern, "NORMAL-A").x == Catch::Approx(entry_with(modern, "BODY-A").x).margin(0.1));
}

TEST_CASE("native PDF places floating pictures in table cells relative to the cell", "[native_pdf]") {
    // One-row table, first column 3000 twips (150 pt) from the margin; the
    // cell's content box starts at 56.7 + 5.4 pt.
    auto table = [](const std::string& cell0, const std::string& cell1 = p("NEXT-CELL")) {
        return R"(<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>)" + cell0 +
               "</w:tc><w:tc>" + cell1 + "</w:tc></w:tr></w:tbl>";
    };
    const double cell_x = 56.7 + 5.4;

    SECTION("a stamp behind the signature: offset from the cell and its paragraph, under the text") {
        const std::string body = table(para_with("SIGNATURE", floating(60, 40, "column", offset(10), "paragraph", offset(5),
                                                                       "<wp:wrapNone/>", /*behind*/ true)));
        const Parts parts = with_picture(body);
        const PageDrawing d = render_drawing(parts);
        REQUIRE(d.images.size() == 1);
        CHECK(d.images[0].x == Catch::Approx(cell_x + 10.0).margin(0.2));
        CHECK(d.images[0].order < d.first_text);
        // Its top 5 pt below the top of the cell's text, which starts at the page's top margin.
        CHECK(d.images[0].y + 40.0 == Catch::Approx(kPageHeight - 56.7 - 5.0).margin(0.5));
        // wrapNone: the row keeps its one line of height.
        const auto entries = render_entries(parts);
        CHECK(entry_with(entries, "SIGNATURE").y == Catch::Approx(entry_with(entries, "NEXT-CELL").y).margin(0.1));
    }
    SECTION("in front, aligned right in the cell") {
        const Parts parts = with_picture(table(para_with("TEXT", floating(30, 20, "column", "<wp:align>right</wp:align>",
                                                                           "paragraph", offset(0), "<wp:wrapNone/>"))));
        const PageDrawing d = render_drawing(parts);
        REQUIRE(d.images.size() == 1);
        CHECK(d.images[0].x + 30.0 == Catch::Approx(56.7 + 150.0 - 5.4).margin(0.2));
        CHECK(d.images[0].order > d.last_text);
    }
    SECTION("layoutInCell=\"0\": positioned on the page") {
        std::string pic = floating(30, 20, "column", offset(300), "paragraph", offset(0), "<wp:wrapNone/>");
        pic.replace(pic.find("layoutInCell=\"1\""), 16, "layoutInCell=\"0\"");
        const PageDrawing d = render_drawing(with_picture(table(para_with("TEXT", pic))));
        REQUIRE(d.images.size() == 1);
        CHECK(d.images[0].x == Catch::Approx(56.7 + 300.0).margin(0.2));
    }
    SECTION("top-and-bottom wrapping pushes the cell's text down and grows the row") {
        const std::string body =
            table(para_with("ANCHOR", floating(40, 100, "column", offset(0), "paragraph", offset(0), "<wp:wrapTopAndBottom/>")) +
                  p("AFTER")) +
            para_with("BELOW-TABLE");
        const auto entries = render_entries(with_picture(body));
        auto y = [&](const char* needle) { return entry_with(entries, needle).y; };
        CHECK(y("NEXT-CELL") - y("ANCHOR") == Catch::Approx(100.0).margin(0.5));  // the anchor's own text right under the picture
        CHECK(y("NEXT-CELL") - y("BELOW-TABLE") > 120.0);
    }
}

TEST_CASE("native PDF draws text boxes with their fill, outline and text", "[native_pdf]") {
    // Yellow box with a 1 pt red outline, 200 x 80 pt at (100, 200) from
    // the page's corner, text centred vertically.
    const std::string box = shape_run(
        true, 100, 200, 200, 80, "rect",
        R"(<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>)",
        {}, p("IN-A-BOX"), R"(<wps:bodyPr anchor="ctr"/>)");
    const Parts parts = make_parts(para_with("ANCHOR-TEXT", box) + p("AFTER"), {}, {}, kSect);
    const PageDrawing d = render_drawing(parts);
    const auto entries = render_entries(parts);

    CHECK(std::count(d.fills.begin(), d.fills.end(), std::string("ffff00")) == 1);
    bool red_outline = false;
    for (const StrokedLine& l : d.lines) red_outline = red_outline || (l.color == "ff0000" && l.width == Catch::Approx(1.0));
    CHECK(red_outline);
    // The box's text inside it, once — not in its anchor paragraph's line.
    const TextEntry& in_box = entry_with(entries, "IN-A-BOX");
    CHECK(in_box.x == Catch::Approx(100.0 + 7.2).margin(0.5));
    CHECK(in_box.y < kPageHeight - 200.0 - 20.0);  // anchored in the middle of the 80 pt box
    CHECK(in_box.y > kPageHeight - 280.0 + 20.0);
    CHECK(std::count_if(entries.begin(), entries.end(), [](const TextEntry& e) { return contains(e.text, "IN-A-BOX"); }) == 1);
    // In front, taking no room: the body flows as if it weren't there.
    CHECK(entry_with(entries, "ANCHOR-TEXT").y - entry_with(entries, "AFTER").y < 20.0);
}

TEST_CASE("native PDF draws inline shapes with their style's colors and fits text to them", "[native_pdf]") {
    // A default Word shape: accent1 fill, darker outline, white text — all
    // from <wps:style>. Auto-fit: 10 pt tall as stored, grown for three lines.
    const std::string style =
        R"(<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>)"
        R"(<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>)"
        R"(<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>)";
    const std::string content = p("WHITE-1") + p("WHITE-2") + "<w:p><w:r><w:rPr><w:color w:val=\"00FF00\"/></w:rPr><w:t>GREEN-3</w:t></w:r></w:p>";
    const std::string shape = shape_run(false, 0, 0, 150, 10, "roundRect", "", style, content,
                                        R"(<wps:bodyPr><a:spAutoFit/></wps:bodyPr>)");
    const Parts parts = make_parts(p("BEFORE") + "<w:p>" + shape + "</w:p>" + p("AFTER"), {}, {}, kSect);
    const PageDrawing d = render_drawing(parts);
    const auto entries = render_entries(parts);
    auto y = [&](const char* needle) { return entry_with(entries, needle).y; };

    // Text colors in drawing order: BEFORE black, the box's two white lines, the green one, AFTER black.
    REQUIRE(d.text_colors.size() == 5);
    CHECK(d.text_colors[0] == "000000");
    CHECK(d.text_colors[1] == "ffffff");
    CHECK(d.text_colors[2] == "ffffff");
    CHECK(d.text_colors[3] == "00ff00");
    CHECK(d.text_colors[4] == "000000");
    // The shape is as tall as its three lines, and the flow continues below it.
    const double line = y("WHITE-1") - y("WHITE-2");
    CHECK(y("BEFORE") - y("AFTER") > 3.0 * line + 7.2);
    CHECK(y("GREEN-3") > y("AFTER"));
}

TEST_CASE("native PDF draws a text box floating in a table cell", "[native_pdf]") {
    std::string box = shape_run(true, 0, 0, 120, 40, "rect", R"(<a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>)",
                                {}, p("CELL-BOX"));
    // Relative to the cell's column and paragraph rather than the page.
    box.replace(box.find(R"(relativeFrom="page")"), 19, R"(relativeFrom="column")");
    box.replace(box.find(R"(relativeFrom="page")"), 19, R"(relativeFrom="paragraph")");
    const std::string body =
        R"(<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>)" + para_with("CELL-TEXT", box) +
        "</w:tc></w:tr></w:tbl>";
    const auto entries = render_entries(make_parts(body, {}, {}, kSect));
    CHECK(entry_with(entries, "CELL-BOX").x == Catch::Approx(56.7 + 5.4 + 7.2).margin(0.5));
    CHECK(entry_with(entries, "CELL-BOX").y < entry_with(entries, "CELL-TEXT").y + 1.0);
}

// ── Footnotes and endnotes ──────────────────────────────────────────────────

constexpr const char* kNoteStyles =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
    R"(<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">)"
    R"(<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>)"
    R"(<w:rPr><w:sz w:val="22"/></w:rPr></w:style>)"
    R"(<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>)"
    R"(<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>)"
    R"(<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>)"
    R"(<w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>)"
    R"(</w:styles>)";

// A footnote/endnote part: Word's separator notes plus `notes` (id, text
// paragraphs) — each note's first paragraph starts with its own mark.
std::string notes_xml(const char* kind, const std::vector<std::pair<int, std::vector<std::string>>>& notes,
                      const std::string& notice = {}) {
    const std::string tag = std::string("w:") + kind;
    std::string x = std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:)") + kind + "s" + kNamespaces + ">";
    x += "<" + tag + R"( w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></)" + tag + ">";
    x += "<" + tag + R"( w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></)" + tag + ">";
    x += "<" + tag + R"( w:type="continuationNotice" w:id="1000"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>)" + notice + "</w:t></w:r></w:p></" + tag + ">";
    for (const auto& [id, paras] : notes) {
        x += "<" + tag + " w:id=\"" + std::to_string(id) + "\">";
        for (std::size_t i = 0; i < paras.size(); ++i) {
            x += R"(<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>)";
            if (i == 0) {
                x += std::string(R"(<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:)") + kind + "Ref/></w:r>"
                     R"(<w:r><w:t xml:space="preserve"> </w:t></w:r>)";
            }
            x += R"(<w:r><w:t xml:space="preserve">)" + paras[i] + "</w:t></w:r></w:p>";
        }
        x += "</" + tag + ">";
    }
    return x + "</w:" + kind + "s>";
}

// A run referencing footnote (or endnote) `id`.
std::string note_ref(int id, const char* kind = "footnote", bool custom = false) {
    return std::string(R"(<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:)") + kind +
           "Reference w:id=\"" + std::to_string(id) + "\"" + (custom ? R"( w:customMarkFollows="1")" : "") +
           "/>" + (custom ? "<w:t>*</w:t>" : "") + "</w:r>";
}

std::string para_with_ref(const std::string& text, int id, const char* kind = "footnote") {
    return R"(<w:p><w:r><w:t xml:space="preserve">)" + text + "</w:t></w:r>" + note_ref(id, kind) + "</w:p>";
}

Parts note_parts(const std::string& body, const std::string& footnotes, const std::string& endnotes = {},
                 const std::string& settings = {}, const std::string& sect = kSect) {
    Parts extra{{"word/styles.xml", kNoteStyles}, {"word/footnotes.xml", footnotes}};
    if (!endnotes.empty()) extra["word/endnotes.xml"] = endnotes;
    if (!settings.empty()) {
        extra["word/settings.xml"] = std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings)") +
                                     kNamespaces + ">" + settings + "</w:settings>";
    }
    return make_parts(body, {{"rIdS", "styles", "styles.xml"}, {"rIdF", "footnotes", "footnotes.xml"},
                             {"rIdE", "endnotes", "endnotes.xml"}, {"rIdT", "settings", "settings.xml"}},
                      extra, sect);
}

// Text entries of page `page` that are exactly `text`.
std::vector<TextEntry> entries_equal(const std::vector<TextEntry>& entries, const std::string& text) {
    std::vector<TextEntry> out;
    for (const auto& e : entries) {
        if (e.text == text) out.push_back(e);
    }
    return out;
}

TEST_CASE("native PDF draws footnotes at the bottom of their page and endnotes at the end", "[native_pdf]") {
    const std::string body =
        para_with_ref("FIRST-REF", 1) +
        R"(<w:p><w:r><w:t xml:space="preserve">CUSTOM-REF</w:t></w:r>)" + note_ref(2, "footnote", true) + "</w:p>" +
        para_with_ref("SECOND-REF", 3) + para_with_ref("END-REF", 1, "endnote") + p("LAST-BODY");
    const auto entries = render_entries(note_parts(
        body, notes_xml("footnote", {{1, {"NOTE-ONE"}}, {2, {"NOTE-CUSTOM"}}, {3, {"NOTE-THREE"}}}),
        notes_xml("endnote", {{1, {"END-TEXT-I"}}})));
    for (const auto& e : entries) INFO(e.page << " " << e.x << "," << e.y << " [" << e.text << "]");

    // All on page 1: footnotes below the body, above the bottom margin, in order.
    const TextEntry& last_body = entry_with(entries, "LAST-BODY");
    const TextEntry& one   = entry_with(entries, "NOTE-ONE");
    const TextEntry& cust  = entry_with(entries, "NOTE-CUSTOM");
    const TextEntry& three = entry_with(entries, "NOTE-THREE");
    CHECK(one.page == 0);
    CHECK(three.page == 0);
    CHECK(one.y < last_body.y);
    CHECK(cust.y < one.y);
    CHECK(three.y < cust.y);
    CHECK(three.y >= 56.7 - 0.5);
    CHECK(one.y < 200.0);  // at the bottom, not after the text
    // The endnote follows the body text.
    const TextEntry& endnote = entry_with(entries, "END-TEXT-I");
    CHECK(endnote.y < last_body.y);
    CHECK(endnote.y > one.y);

    // Marks: 1, *, 2 (the custom mark takes no number); the endnote is "i".
    CHECK(entries_equal(entries, "1").size() == 2);  // reference + the note's own mark
    CHECK(entries_equal(entries, "2").size() == 2);
    CHECK(entries_equal(entries, "3").empty());
    CHECK(entries_equal(entries, "i").size() == 2);
    // The reference is raised above its line's baseline.
    const TextEntry& first = entry_with(entries, "FIRST-REF");
    bool raised = false;
    for (const auto& e : entries_equal(entries, "1")) {
        if (std::abs(e.y - first.y) < 8.0) raised = e.y > first.y + 1.0;
    }
    CHECK(raised);
}

TEST_CASE("native PDF keeps every footnote on the page of its reference", "[native_pdf]") {
    std::string body;
    std::vector<std::pair<int, std::vector<std::string>>> notes;
    for (int i = 1; i <= 70; ++i) {
        body += para_with_ref("PARA-" + std::to_string(i) + "-X", i);
        notes.push_back({i, {"NOTE-" + std::to_string(i) + "-X"}});
    }
    const auto entries = render_entries(note_parts(body, notes_xml("footnote", notes)));
    int pages = 0;
    for (int i = 1; i <= 70; ++i) {
        const TextEntry& para = entry_with(entries, "PARA-" + std::to_string(i) + "-X");
        const TextEntry& note = entry_with(entries, "NOTE-" + std::to_string(i) + "-X");
        INFO("note " << i);
        CHECK(para.page == note.page);
        CHECK(note.y >= 56.7 - 0.5);
        // Below every body line of its page.
        for (const auto& e : entries) {
            if (e.page == note.page && e.text.rfind("PARA-", 0) == 0) CHECK(e.y > note.y);
        }
        pages = std::max(pages, para.page + 1);
    }
    CHECK(pages >= 2);
}

TEST_CASE("native PDF continues a long footnote on the next page", "[native_pdf]") {
    std::vector<std::string> lines;
    for (int i = 1; i <= 90; ++i) lines.push_back("LONG-" + std::to_string(i) + "-X");
    const std::string body = para_with_ref("REF-PARA", 1) + p("AFTER-PARA");
    const auto entries = render_entries(note_parts(body, notes_xml("footnote", {{1, lines}})));

    const TextEntry& ref = entry_with(entries, "REF-PARA");
    CHECK(ref.page == 0);
    CHECK(entry_with(entries, "LONG-1-X").page == 0);
    // Each line once, in order, never below the bottom margin or overlapping body text.
    int last_page = 0;
    double last_y = 1e9;
    for (int i = 1; i <= 90; ++i) {
        const std::string needle = "LONG-" + std::to_string(i) + "-X";
        int count = 0;
        for (const auto& e : entries) count += e.text.find(needle) != std::string::npos ? 1 : 0;
        INFO(needle);
        CHECK(count == 1);
        const TextEntry& e = entry_with(entries, needle);
        CHECK(e.y >= 56.7 - 0.5);
        CHECK(e.y <= 841.9 - 56.7);
        CHECK((e.page > last_page || e.y < last_y));
        last_page = e.page;
        last_y = e.y;
    }
    CHECK(last_page >= 1);
    // The body goes on above the continued note.
    const TextEntry& after = entry_with(entries, "AFTER-PARA");
    for (const auto& e : entries) {
        if (e.page == after.page && e.text.rfind("LONG-", 0) == 0) CHECK(e.y < after.y);
    }
}

TEST_CASE("native PDF numbers footnotes as footnotePr says", "[native_pdf]") {
    const std::string notes = notes_xml("footnote", {{1, {"NOTE-A"}}, {2, {"NOTE-B"}}});
    const std::string body = para_with_ref("PAGE-ONE", 1) +
                             R"(<w:p><w:r><w:br w:type="page"/></w:r></w:p>)" + para_with_ref("PAGE-TWO", 2);
    auto marks = [&](const std::string& settings) {
        const auto entries = render_entries(note_parts(body, notes, {}, settings));
        std::vector<std::string> out;
        for (const auto& e : entries) {
            if (e.text != "PAGE-ONE" && e.text != "PAGE-TWO" && e.text.find("NOTE-") == std::string::npos &&
                e.text.find_first_not_of(" ") != std::string::npos) {
                out.push_back(std::to_string(e.page) + ":" + e.text);
            }
        }
        return out;
    };
    using V = std::vector<std::string>;
    CHECK(marks("") == V{"0:1", "0:1", "1:2", "1:2"});
    CHECK(marks(R"(<w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>)") ==
          V{"0:1", "0:1", "1:1", "1:1"});
    CHECK(marks(R"(<w:footnotePr><w:numFmt w:val="chicago"/></w:footnotePr>)") ==
          V{"0:*", "0:*", "1:†", "1:†"});
    CHECK(marks(R"(<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/></w:footnotePr>)") ==
          V{"0:iii", "0:iii", "1:iv", "1:iv"});
}

TEST_CASE("native PDF places footnotes referenced in table cells", "[native_pdf]") {
    const std::string body =
        R"(<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>)"
        R"(<w:tr><w:tc>)" + para_with_ref("CELL-REF", 1) + "</w:tc></w:tr></w:tbl>" + p("BELOW-TABLE");
    const auto entries = render_entries(note_parts(body, notes_xml("footnote", {{1, {"CELL-NOTE"}}})));
    const TextEntry& note = entry_with(entries, "CELL-NOTE");
    CHECK(note.page == 0);
    CHECK(note.y < entry_with(entries, "BELOW-TABLE").y);
    CHECK(note.y < 200.0);
}

TEST_CASE("native PDF puts beneathText footnotes right under the text", "[native_pdf]") {
    const std::string body = para_with_ref("SHORT-BODY", 1);
    const std::string notes = notes_xml("footnote", {{1, {"UNDER-NOTE"}}});
    const auto bottom = render_entries(note_parts(body, notes));
    const auto beneath = render_entries(
        note_parts(body, notes, {}, R"(<w:footnotePr><w:pos w:val="beneathText"/></w:footnotePr>)"));
    CHECK(entry_with(bottom, "UNDER-NOTE").y < 200.0);
    const double text_y = entry_with(beneath, "SHORT-BODY").y;
    const double note_y = entry_with(beneath, "UNDER-NOTE").y;
    CHECK(note_y < text_y);
    CHECK(note_y > text_y - 50.0);  // the separator and the note's line, nothing more
}

TEST_CASE("native PDF prints the continuation notice under a footnote that goes on", "[native_pdf]") {
    std::vector<std::string> lines;
    for (int i = 1; i <= 90; ++i) lines.push_back("CONT-" + std::to_string(i) + "-X");
    const auto entries = render_entries(note_parts(para_with_ref("REF-PARA", 1) + p("AFTER-PARA"),
                                                   notes_xml("footnote", {{1, lines}}, "NOTICE-TEXT")));
    std::vector<TextEntry> notices;
    for (const auto& e : entries) {
        if (e.text.find("NOTICE-TEXT") != std::string::npos) notices.push_back(e);
    }
    REQUIRE(!notices.empty());
    int last_page = 0;
    for (const auto& e : entries) last_page = std::max(last_page, e.page);
    for (const auto& n : notices) {
        CHECK(n.y >= 56.7 - 0.5);
        CHECK(n.page < entry_with(entries, "CONT-90-X").page);  // not under the note's end
        // Under the last line of the note on its page.
        for (const auto& e : entries) {
            if (e.page == n.page && e.text.rfind("CONT-", 0) == 0) CHECK(e.y > n.y);
        }
    }
    for (int i = 1; i <= 90; ++i) {
        int count = 0;
        for (const auto& e : entries) count += e.text.find("CONT-" + std::to_string(i) + "-X") != std::string::npos;
        CHECK(count == 1);
    }
    CHECK(last_page >= 1);
}

TEST_CASE("native PDF keeps a heading with the next paragraph and its footnote", "[native_pdf]") {
    // The filler leaves ~42 pt: room for the heading and the next line, not
    // for them plus the next line's footnote (and its separator).
    const std::string filler = p("FILLER", R"(<w:spacing w:before="13500" w:after="0"/>)");
    const std::string heading = p("THE-HEADING", R"(<w:keepNext/><w:spacing w:after="0"/>)");
    std::vector<std::string> lines;
    for (int i = 1; i <= 6; ++i) lines.push_back("KN-NOTE-" + std::to_string(i));
    const std::string notes = notes_xml("footnote", {{1, lines}});

    const auto without = render_entries(note_parts(filler + heading + p("PLAIN-NEXT"), notes));
    CHECK(entry_with(without, "THE-HEADING").page == 0);
    CHECK(entry_with(without, "PLAIN-NEXT").page == 0);

    const auto with = render_entries(note_parts(filler + heading + para_with_ref("NOTED-NEXT", 1), notes));
    CHECK(entry_with(with, "NOTED-NEXT").page == 1);
    CHECK(entry_with(with, "THE-HEADING").page == 1);
    CHECK(entry_with(with, "KN-NOTE-1").page == 1);
}

TEST_CASE("native PDF continues a footnote into a section with other margins", "[native_pdf]") {
    // Paragraphs that wrap differently at the two text widths.
    std::vector<std::string> paras;
    for (int i = 1; i <= 40; ++i) {
        paras.push_back("WIDE-" + std::to_string(i) + "-X lorem ipsum dolor sit amet consectetur adipiscing elit "
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua");
    }
    const std::string narrow_sect =
        R"(<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>)"
        R"(<w:pgMar w:top="1134" w:right="3402" w:bottom="1134" w:left="1134"/></w:sectPr></w:pPr></w:p>)";
    const std::string body = para_with_ref("REF-PARA", 1) + narrow_sect + p("SECOND-SECTION");
    const auto entries = render_entries(note_parts(body, notes_xml("footnote", {{1, paras}})));
    int last_page = 0;
    double last_y = 1e9;
    for (int i = 1; i <= 40; ++i) {
        const std::string needle = "WIDE-" + std::to_string(i) + "-X";
        int count = 0;
        for (const auto& e : entries) count += e.text.find(needle) != std::string::npos;
        INFO(needle);
        CHECK(count == 1);
        const TextEntry& e = entry_with(entries, needle);
        CHECK((e.page > last_page || e.y < last_y));
        last_page = e.page;
        last_y = e.y;
    }
    CHECK(last_page >= 1);
}


TEST_CASE("native PDF flows text through columns", "[native_pdf]") {
    constexpr double kLeft = 56.7, kWidth = 595.3 - 2 * 56.7;
    // kSect with columns; `type`: this section's start.
    auto sect = [](const std::string& cols, const std::string& type = {}) {
        return R"(<w:sectPr>)" + (type.empty() ? std::string() : R"(<w:type w:val=")" + type + R"("/>)") +
               R"(<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>)" +
               cols + "</w:sectPr>";
    };
    auto paras = [](const std::string& prefix, int n) {
        std::string out;
        for (int i = 0; i < n; ++i) out += para_with(prefix + std::to_string(i));
        return out;
    };
    const double col2 = kLeft + (kWidth - 36.0) / 2.0 + 36.0;  // default space: 1/2 inch

    SECTION("a full column continues at the top of the next, then on the next page") {
        const auto entries = render_entries(make_parts(paras("L-", 140), {}, {}, sect(R"(<w:cols w:num="2"/>)")));
        const TextEntry& first = entry_with(entries, "L-0");
        CHECK(first.x == Catch::Approx(kLeft).margin(0.5));
        // The first line of the right column: at the left column's top.
        const auto right = std::find_if(entries.begin(), entries.end(), [&](const TextEntry& e) {
            return e.page == 0 && e.x > kLeft + 100.0;
        });
        REQUIRE(right != entries.end());
        CHECK(right->x == Catch::Approx(col2).margin(0.5));
        CHECK(right->y == Catch::Approx(first.y).margin(0.5));
        // Lines keep their order: left column, right column, next page.
        int last = -1, page = 0;
        bool left_side = true;
        for (const auto& e : entries) {
            const int i = std::stoi(e.text.substr(2));
            INFO(e.text << " p" << e.page << " x" << e.x);
            CHECK(i == last + 1);
            last = i;
            const bool now_left = e.x < kLeft + 100.0;
            if (e.page == page && !left_side) CHECK_FALSE(now_left);  // no going back
            if (e.page != page) { page = e.page; left_side = true; }
            left_side = left_side && now_left;
        }
        CHECK(page >= 1);
    }
    SECTION("a column break starts the next column") {
        const auto entries = render_entries(make_parts(
            para_with("TOP-LEFT") + R"(<w:p><w:r><w:br w:type="column"/><w:t>TOP-RIGHT</w:t></w:r></w:p>)",
            {}, {}, sect(R"(<w:cols w:num="2"/>)")));
        CHECK(entry_with(entries, "TOP-RIGHT").x == Catch::Approx(col2).margin(0.5));
        CHECK(entry_with(entries, "TOP-RIGHT").y == Catch::Approx(entry_with(entries, "TOP-LEFT").y).margin(0.5));
        CHECK(entry_with(entries, "TOP-RIGHT").page == 0);
    }
    SECTION("unequal columns: each <w:col>'s width and space") {
        const auto entries = render_entries(make_parts(
            para_with("NARROW") + R"(<w:p><w:r><w:br w:type="column"/><w:t>WIDE</w:t></w:r></w:p>)", {}, {},
            sect(R"(<w:cols w:num="2" w:equalWidth="0"><w:col w:w="2000" w:space="360"/><w:col w:w="7638"/></w:cols>)")));
        CHECK(entry_with(entries, "WIDE").x == Catch::Approx(kLeft + 100.0 + 18.0).margin(0.5));
    }
    SECTION("columns before a continuous section break are levelled, the text after goes below") {
        const std::string body =
            para_with("BEFORE") +
            "<w:p><w:pPr>" + sect("") + "</w:pPr></w:p>" +                      // single column
            paras("C-", 9) + "<w:p><w:pPr>" + sect(R"(<w:cols w:num="3" w:sep="1"/>)", "continuous") +
            "</w:pPr></w:p>" + para_with("AFTER");
        const auto entries = render_entries(make_parts(body, {}, {}, sect("", "continuous")));
        std::map<int, int> per_column;
        double lowest = 1e9, top = -1e9;
        for (const auto& e : entries) {
            if (e.text.rfind("C-", 0) != 0) continue;
            INFO(e.text << " x" << e.x << " y" << e.y);
            CHECK(e.page == 0);
            ++per_column[static_cast<int>((e.x - kLeft) / (kWidth / 3.0))];
            lowest = std::min(lowest, e.y);
            top = std::max(top, e.y);
        }
        // 9 + the empty paragraph holding the section break: 10 lines,
        // 4 + 4 + 2 (one of them that empty paragraph), as in Word.
        REQUIRE(per_column.size() == 3);
        CHECK(per_column[0] == 4);
        CHECK(per_column[1] == 4);
        CHECK(per_column[2] == 1);
        CHECK(top < entry_with(entries, "BEFORE").y);
        const TextEntry& after = entry_with(entries, "AFTER");
        CHECK(after.page == 0);
        CHECK(after.x == Catch::Approx(kLeft).margin(0.5));
        CHECK(after.y < lowest);
        CHECK(after.y > lowest - 40.0);
    }
}

TEST_CASE("native PDF replaces a font without Cyrillic by one of the same kind", "[native_pdf]") {
    // The fonts embedded in the PDF of `body`, by their base names.
    auto fonts_of = [](const std::string& body) {
        const Parts parts = make_parts(body);
        pugi::xml_document doc;
        const std::string& xml = parts.at("word/document.xml");
        REQUIRE(doc.load_buffer(xml.data(), xml.size(), docweft::docx::kXmlParse));
        const fs::path out = temp_pdf("docweft_native_fonts");
        docweft::docx::render_native_pdf(doc, parts, out);
        PoDoFo::PdfMemDocument pdf;
        pdf.Load(out.string());
        std::string names;
        for (const PoDoFo::PdfObject* obj : pdf.GetObjects()) {
            const PoDoFo::PdfDictionary* dict = nullptr;
            if (obj->TryGetDictionary(dict) && dict->HasKey("BaseFont") && dict->HasKey("Subtype") &&
                dict->MustFindKey("Subtype").GetName() == "Type0") {
                names += std::string(dict->MustFindKey("BaseFont").GetName().GetString()) + " ";
            }
        }
        fs::remove(out);
        return names;
    };
    auto run = [](const char* font, const char* rpr, const char* text) {
        return std::string(R"(<w:p><w:r><w:rPr><w:rFonts w:ascii=")") + font + R"(" w:hAnsi=")" + font + R"("/>)" +
               rpr + "</w:rPr><w:t>" + text + "</w:t></w:r></w:p>";
    };
    // Needs Caladea (Cambria's substitute, Latin only) and DejaVu Serif /
    // Sans Mono, the first serif and monospaced replacements tried.
    PoDoFo::PdfMemDocument probe;
    for (const char* family : {"Caladea", "DejaVu Serif", "DejaVu Sans Mono"}) {
        PoDoFo::PdfFont* f = probe.GetFonts().SearchFont(family);
        if (f == nullptr || std::string(f->GetMetrics().GeFontFamilyNameSafe()) != family) {
            SKIP(std::string(family) + " isn't installed");
        }
    }

    const std::string serif_fonts = fonts_of(run("Cambria", "", "Привет") + run("Cambria", "<w:i/>", "мир"));
    INFO(serif_fonts);
    CHECK(contains(serif_fonts, "DejaVuSerif"));
    CHECK(contains(serif_fonts, "Italic"));
    CHECK_FALSE(contains(serif_fonts, "DejaVuSans "));
    CHECK_FALSE(contains(serif_fonts, "DejaVuSans-"));

    // Only the letters Caladea lacks are replaced: the comma stays in it.
    const std::string mixed_fonts = fonts_of(run("Cambria", "", "выросла,"));
    INFO(mixed_fonts);
    CHECK(contains(mixed_fonts, "DejaVuSerif"));
    CHECK(contains(mixed_fonts, "Caladea"));

    const std::string mono_fonts = fonts_of(run("Consolas", "", "Код"));
    INFO(mono_fonts);
    CHECK(contains(mono_fonts, "DejaVuSansMono"));
}

TEST_CASE("native PDF spaces paragraphs by the larger of space after and before", "[native_pdf]") {
    // Exact 12 pt lines; the first paragraph has 10 pt after, the second 24 pt before.
    const std::string body =
        R"(<w:p><w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>FIRST</w:t></w:r></w:p>)"
        R"(<w:p><w:pPr><w:spacing w:before="480" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>SECOND</w:t></w:r></w:p>)";
    auto gap = [&](const std::string& compat) {
        Parts extra;
        if (!compat.empty()) {
            extra["word/settings.xml"] = std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings)") +
                                         kNamespaces + "><w:compat>" + compat + "</w:compat></w:settings>";
        }
        const auto entries = render_entries(make_parts(body, {}, extra));
        return entry_with(entries, "FIRST").y - entry_with(entries, "SECOND").y;
    };
    CHECK(gap("") == Catch::Approx(12.0 + 24.0).margin(0.1));
    // "Don't use HTML paragraph auto spacing": they add up.
    CHECK(gap("<w:doNotUseHTMLParagraphAutoSpacing/>") == Catch::Approx(12.0 + 10.0 + 24.0).margin(0.1));
}

TEST_CASE("native PDF keeps widow and orphan lines off page edges", "[native_pdf]") {
    // Exact 12 pt lines, no spacing: kSect's body holds 60 of them.
    const std::string ppr = R"(<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>)";
    auto filler = [&](int n) {
        std::string out;
        for (int i = 0; i < n; ++i) out += "<w:p><w:pPr>" + ppr + "</w:pPr><w:r><w:t>F" + std::to_string(i) + "</w:t></w:r></w:p>";
        return out;
    };
    auto para = [&](int lines, const std::string& extra = {}) {
        std::string runs;
        for (int i = 1; i <= lines; ++i) runs += (i > 1 ? "<w:br/>" : "") + std::string("<w:t>L") + std::to_string(i) + "</w:t>";
        return "<w:p><w:pPr>" + ppr + extra + "</w:pPr><w:r>" + runs + "</w:r></w:p>";
    };
    auto page_of = [](const std::vector<TextEntry>& entries, const std::string& text) {
        for (const auto& e : entries) if (e.text == text) return e.page;
        return -1;
    };
    SECTION("the last line doesn't go over alone: the one before it goes along") {
        const auto entries = render_entries(make_parts(filler(57) + para(4), {}, {}, kSect));
        CHECK(page_of(entries, "L2") == 0);
        CHECK(page_of(entries, "L3") == 1);
        CHECK(page_of(entries, "L4") == 1);
    }
    SECTION("widowControl off: it does") {
        const auto entries = render_entries(
            make_parts(filler(57) + para(4, R"(<w:widowControl w:val="0"/>)"), {}, {}, kSect));
        CHECK(page_of(entries, "L3") == 0);
        CHECK(page_of(entries, "L4") == 1);
    }
    SECTION("the first line doesn't stay alone: the paragraph moves") {
        const auto entries = render_entries(make_parts(filler(59) + para(3), {}, {}, kSect));
        CHECK(page_of(entries, "F58") == 0);
        CHECK(page_of(entries, "L1") == 1);
    }
    SECTION("three lines with room for two: all move, none left alone") {
        const auto entries = render_entries(make_parts(filler(58) + para(3), {}, {}, kSect));
        CHECK(page_of(entries, "L1") == 1);
        CHECK(page_of(entries, "L3") == 1);
    }
}

TEST_CASE("native PDF gives table borders room in the row height", "[native_pdf]") {
    // Two-row tables with all borders `sz` eighths of a point; exact 20 pt
    // lines (above a row's one-line minimum), no cell margins.
    auto row_pitch = [](int sz, double* first_top = nullptr) {
        const std::string b = R"( w:val="single" w:sz=")" + std::to_string(sz) + R"(" w:space="0" w:color="000000"/>)";
        const std::string ppr = R"(<w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr>)";
        auto row = [&](const char* text) {
            return std::string("<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p>") + ppr +
                   "<w:r><w:t>" + text + "</w:t></w:r></w:p></w:tc></w:tr>";
        };
        const std::string body =
            R"(<w:tbl><w:tblPr><w:tblBorders><w:top)" + b + "<w:left" + b + "<w:bottom" + b + "<w:right" + b +
            "<w:insideH" + b + "<w:insideV" + b + R"(</w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/>)"
            R"(<w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>)" +
            row("ROWA") + row("ROWB") + "</w:tbl>" + p("AFTER");
        const auto entries = render_entries(make_parts(body, {}, {}, kSect));
        if (first_top != nullptr) *first_top = entry_with(entries, "ROWA").y;
        return std::pair{entry_with(entries, "ROWA").y - entry_with(entries, "ROWB").y,
                         entry_with(entries, "ROWB").y - entry_with(entries, "AFTER").y};
    };
    double top_thin = 0.0, top_thick = 0.0;
    const auto [thin_rows, thin_after] = row_pitch(4, &top_thin);     // 0.5 pt
    const auto [thick_rows, thick_after] = row_pitch(24, &top_thick);  // 3 pt
    // Between rows: the line plus one border.
    CHECK(thin_rows == Catch::Approx(20.0 + 0.5).margin(0.05));
    CHECK(thick_rows == Catch::Approx(20.0 + 3.0).margin(0.05));
    // The text sits under the table's top border, the next paragraph under its bottom one.
    CHECK(top_thin - top_thick == Catch::Approx(2.5).margin(0.05));
    CHECK(thick_after - thin_after == Catch::Approx(2.5).margin(0.05));
}

TEST_CASE("native PDF lays out chart axes as LibreOffice does", "[native_pdf]") {
    const std::string bars = R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)";
    SECTION("scale: 5 % headroom, the finest 1/2/5 step that fits the axis' length") {
        // 3400 → 3570: steps 100 and 200 would crowd a ~108 pt axis; 500 gives 8 intervals.
        const auto entries = render_chart(bars + series_xml(0, "A", 1700, 3400) +
                                          R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes);
        CHECK(has_entry(entries, "500"));
        CHECK(has_entry(entries, "4000"));
        CHECK_FALSE(has_entry(entries, "4500"));
    }
    SECTION("a value axis title stands turned, left of the value labels") {
        const auto entries = render_chart(
            bars + series_xml(0, "A", 3, 5) + R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)"
            R"(<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>)"
            R"(<c:valAx><c:axId val="2"/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>VALTITLE</a:t></a:r></a:p>)"
            R"(</c:rich></c:tx></c:title><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>)");
        REQUIRE(has_entry(entries, "VALTITLE"));
        CHECK(entry_with(entries, "VALTITLE").x < entry_with(entries, "Q1").x);
        // Turned: it starts within the axis' height (above its "0", below
        // the plot's top) rather than sitting over the plot.
        const double zero = entry_with(entries, "0").y;
        CHECK(entry_with(entries, "VALTITLE").y > zero);
        CHECK(entry_with(entries, "VALTITLE").y < zero + 100.0);
    }
    SECTION("area charts put their points on the category ticks, edge to edge") {
        auto spread = [&](const std::string& plot) {
            const auto entries = render_chart(plot + kAxes);
            return entry_with(entries, "Q2").x - entry_with(entries, "Q1").x;
        };
        const double area = spread(R"(<c:areaChart><c:grouping val="standard"/>)" + series_xml(0, "A", 3, 5) +
                                   R"(<c:axId val="1"/><c:axId val="2"/></c:areaChart>)");
        const double bar = spread(bars + series_xml(0, "A", 3, 5) + R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)");
        CHECK(area > bar * 1.5);  // the whole plot width, not half of it
    }
}

TEST_CASE("native PDF wraps long chart titles", "[native_pdf]") {
    // 18 pt bold (the default) in the 236 pt wide test chart: a line may
    // take 80 % of it.
    auto title_lines = [](const std::string& rich_paragraphs) {
        const std::string chart =
            std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") + kNamespaces +
            "><c:chart><c:title><c:tx><c:rich><a:bodyPr/>" + rich_paragraphs + "</c:rich></c:tx></c:title><c:plotArea>"
            R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)" + series_xml(0, "A", 3, 5) +
            R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)" + kAxes + "</c:plotArea></c:chart></c:chartSpace>";
        const auto entries = render_entries(make_parts(chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
                                                       {{"word/charts/chart1.xml", chart}}, kSect));
        std::set<long> ys;
        for (const auto& e : entries) {
            if (e.text.find("WORD") != std::string::npos) ys.insert(std::lround(e.y));
        }
        return ys.size();
    };
    auto para = [](const std::string& text) { return "<a:p><a:r><a:t>" + text + "</a:t></a:r></a:p>"; };
    CHECK(title_lines(para("WORDA WORDB")) == 1);
    CHECK(title_lines(para("WORDA WORDB WORDC WORDD WORDE WORDF WORDG WORDH")) >= 2);
    // Each paragraph of the title's rich text starts a line of its own.
    CHECK(title_lines(para("WORDA") + para("WORDB")) == 2);
}

TEST_CASE("native PDF wraps long axis titles", "[native_pdf]") {
    // Bold 10 pt axis titles in the 236 x 142 pt test chart.
    auto axis_title_lines = [](const std::string& cat_title, const std::string& val_title) {
        auto title = [](const std::string& text) {
            return "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>" + text + "</a:t></a:r></a:p></c:rich></c:tx></c:title>";
        };
        const auto entries = render_chart(
            R"(<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>)" + series_xml(0, "A", 3, 5) +
            R"(<c:axId val="1"/><c:axId val="2"/></c:barChart>)"
            R"(<c:catAx><c:axId val="1"/>)" + title(cat_title) + R"(<c:axPos val="b"/><c:crossAx val="2"/></c:catAx>)"
            R"(<c:valAx><c:axId val="2"/>)" + title(val_title) + R"(<c:axPos val="l"/><c:crossAx val="1"/></c:valAx>)");
        std::set<long> cat_rows, val_columns;
        for (const auto& e : entries) {
            if (e.text.find("CAT") != std::string::npos) cat_rows.insert(std::lround(e.y));
            if (e.text.find("VAL") != std::string::npos) val_columns.insert(std::lround(e.x));
        }
        return std::pair{cat_rows.size(), val_columns.size()};
    };
    const auto [short_cat, short_val] = axis_title_lines("CATA CATB", "VALA VALB");
    CHECK(short_cat == 1);
    CHECK(short_val == 1);
    // Wider than 80 % of the chart / longer than the plot box's height.
    const auto [long_cat, long_val] = axis_title_lines(
        "CATA CATB CATC CATD CATE CATF CATG CATH CATI CATJ CATK CATL CATM CATN",
        "VALA VALB VALC VALD VALE VALF VALG VALH VALI VALJ VALK VALL VALM VALN");
    CHECK(long_cat >= 2);   // rows, one under another
    CHECK(long_val >= 2);   // turned: columns side by side
}

#endif  // DOCWEFT_HAVE_PODOFO
