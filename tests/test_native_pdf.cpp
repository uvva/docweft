// Tests for the native PoDoFo PDF renderer (docx/pdf_native.cpp): content
// that must not be dropped, and strict mode rejecting what it can't draw.
// Only built into the test binary when DOCWEFT_ENABLE_NATIVE_PDF is on.

#if defined(DOCWEFT_HAVE_PODOFO)

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include <podofo/podofo.h>
#include <pugixml.hpp>

#include "docweft/error.hpp"
#include "docx/pdf_native.hpp"
#include "fixture_helpers.hpp"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <sstream>
#include <cstdlib>
#include <filesystem>
#include <string>
#include <unordered_map>
#include <vector>

namespace fs = std::filesystem;

namespace {

using Parts = std::unordered_map<std::string, std::string>;

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
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));

    const fs::path out = fs::temp_directory_path() / "docweft_native_test.pdf";
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
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));

    const fs::path out = fs::temp_directory_path() / "docweft_native_test_entries.pdf";
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
std::string chart_with_plot(const std::string& plot) {
    return std::string(R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace)") + kNamespaces +
           "><c:chart><c:plotArea>" + plot + "</c:plotArea></c:chart></c:chartSpace>";
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

std::vector<TextEntry> render_chart(const std::string& plot) {
    return render_entries(make_parts(chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
                                     {{"word/charts/chart1.xml", chart_with_plot(plot)}}, kSect));
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
        REQUIRE(has_entry(entries, "40000"));  // sum 31080; clustered would stop at 25000
        REQUIRE_FALSE(has_entry(entries, "25000"));
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
        R"(<c:valAx><c:axId val="4"/><c:axPos val="r"/><c:crossAx val="3"/></c:valAx>)");
    REQUIRE(has_entry(entries, "25000"));  // primary axis, left
    REQUIRE(has_entry(entries, "30"));     // secondary axis, right
    REQUIRE(entry_with(entries, "30").x > entry_with(entries, "25000").x + 100.0);
    REQUIRE(has_entry(entries, "Revenue"));
    REQUIRE(has_entry(entries, "Margin"));
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
    const fs::path out = fs::temp_directory_path() / "docweft_native_sections.pdf";
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
    const std::string footnote =
        R"(<w:p><w:r><w:t>With a note</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>)";
    const std::string columns =
        R"(<w:sectPr><w:cols w:num="2"/></w:sectPr>)";

    SECTION("strict (default): NotImplemented listing every problem") {
        const auto parts = make_parts(
            footnote + chart_paragraph("rIdC"), {{"rIdC", "chart", "charts/chart1.xml"}},
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
            REQUIRE(contains(msg, "footnotes"));
            REQUIRE(contains(msg, "multi-column"));
            REQUIRE(contains(msg, "pie charts combined with other chart types"));
            REQUIRE(contains(msg, "DOCWEFT_NATIVE_PDF_STRICT"));
        }
    }
    SECTION("DOCWEFT_NATIVE_PDF_STRICT=0 renders anyway") {
        ScopedEnvVar lenient("DOCWEFT_NATIVE_PDF_STRICT", "0");
        REQUIRE(contains(all_text(render_pages(make_parts(footnote, {}, {}, columns))), "With a note"));
    }
}

TEST_CASE("native PDF strict mode rejects text boxes and unsupported images", "[native_pdf]") {
    const std::string textbox =
        R"(<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="1000" cy="1000"/><a:graphic>)"
        R"(<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">)"
        R"(<w:txbxContent>)" + p("in a box") + R"(</w:txbxContent></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>)";
    const std::string emf_picture =
        R"(<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1000" cy="1000"/><a:graphic>)"
        R"(<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill>)"
        R"(<a:blip r:embed="rIdE"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>)";
    std::string emf(44, '\0');
    emf.replace(40, 4, " EMF");

    try {
        render_pages(make_parts(textbox + emf_picture, {{"rIdE", "image", "media/image1.emf"}},
                                {{"word/media/image1.emf", emf}}));
        FAIL("expected NotImplemented");
    } catch (const docweft::ReportException& e) {
        REQUIRE(e.code() == docweft::ReportError::NotImplemented);
        const std::string msg = e.what();
        REQUIRE(contains(msg, "text boxes"));
        REQUIRE(contains(msg, "EMF"));
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
    int first_text = -1, last_text = -1;  // order of the first/last text shown
};

// Renders `parts` and reads back what page `page`'s content stream draws.
PageDrawing render_drawing(const Parts& parts, unsigned page = 0) {
    pugi::xml_document doc;
    const std::string& xml = parts.at("word/document.xml");
    REQUIRE(doc.load_buffer(xml.data(), xml.size()));
    const fs::path out = fs::temp_directory_path() / "docweft_native_test_drawing.pdf";
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
    body = para_with("FIRST", floating(100, 80, "page", offset(200), "page", offset(300), "<wp:wrapSquare wrapText=\"bothSides\"/>")) + body;
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
        REQUIRE(doc.load_buffer(xml.data(), xml.size()));
        const fs::path out = fs::temp_directory_path() / "docweft_native_test_hidden_image.pdf";
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

#endif  // DOCWEFT_HAVE_PODOFO
