// native_pdf_report — same kind of report as basic_report, but demonstrates
// DocWeft's native PoDoFo PDF backend (DOCWEFT_ENABLE_NATIVE_PDF): the
// resulting PDF is produced with no Microsoft Word or LibreOffice involved
// at all, not even installed on the machine. Includes a bar chart and a pie
// chart to show the backend's chart support (bar/line/area/pie/doughnut —
// see PLAN.md for the full scope and known gaps, e.g. no pie-of-pie/radar/
// scatter/bubble/stock/surface yet).
//
// This binary only exists when the library was built with
// -DDOCWEFT_ENABLE_NATIVE_PDF=ON (see examples/CMakeLists.txt) — save()
// would otherwise throw NoConverter once Word/LibreOffice are disabled
// below, which is the whole point of this example.
//
// Usage:
//   ./native_pdf_report [output.pdf]    # default: native_report.pdf

#include <docweft/error.hpp>
#include <docweft/merger.hpp>

#include <zip.h>

#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>
#include <string_view>
#include <vector>

namespace fs = std::filesystem;

namespace {

struct Part {
    std::string name;
    std::string data;
};

int write_docx(const fs::path& out, const std::vector<Part>& parts) {
    std::error_code ec;
    fs::remove(out, ec);

    int err = 0;
    zip_t* z = zip_open(out.string().c_str(), ZIP_CREATE | ZIP_EXCL, &err);
    if (!z) {
        std::cerr << "zip_open failed (err=" << err << ")\n";
        return 1;
    }
    for (const auto& p : parts) {
        zip_source_t* src = zip_source_buffer(z, p.data.data(), p.data.size(), 0);
        if (!src || zip_file_add(z, p.name.c_str(), src,
                                  ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
            std::cerr << "zip_file_add failed for " << p.name << "\n";
            if (src) zip_source_free(src);
            zip_discard(z);
            return 1;
        }
    }
    if (zip_close(z) != 0) {
        std::cerr << "zip_close failed: " << zip_strerror(z) << "\n";
        zip_discard(z);
        return 1;
    }
    return 0;
}

// Same boilerplate parts as generate_template.cpp, plus a second chart
// (pie) alongside the bar chart to show off the fuller range of chart
// types the native PDF backend now renders.

constexpr std::string_view kContentTypes = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/word/charts/chart2.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>)";

constexpr std::string_view kPackageRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>)";

constexpr std::string_view kDocumentRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rIdChart1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                Target="charts/chart1.xml"/>
  <Relationship Id="rIdChart2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                Target="charts/chart2.xml"/>
</Relationships>)";

// Bar chart: one series × three categories — same shape as
// generate_template.cpp's, kept here so this example is self-contained.
constexpr std::string_view kBarChart = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/>
    <c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Measurements</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>Region A</c:v></c:pt><c:pt idx="1"><c:v>Region B</c:v></c:pt><c:pt idx="2"><c:v>Region C</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>42</c:v></c:pt><c:pt idx="1"><c:v>73.5</c:v></c:pt><c:pt idx="2"><c:v>55</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="111111111"/><c:axId val="222222222"/>
    </c:barChart>
    <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
    <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>)";

// Pie chart: one series, four categories, <c:varyColors val="1"/> so the
// native backend colors per-category (matching Word/Excel's own convention
// for pie/doughnut) instead of per-series.
constexpr std::string_view kPieChart = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>
          <c:pt idx="0"><c:v>Chrome</c:v></c:pt><c:pt idx="1"><c:v>Firefox</c:v></c:pt>
          <c:pt idx="2"><c:v>Safari</c:v></c:pt><c:pt idx="3"><c:v>Other</c:v></c:pt>
        </c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/>
          <c:pt idx="0"><c:v>50</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt>
          <c:pt idx="2"><c:v>25</c:v></c:pt><c:pt idx="3"><c:v>10</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea></c:chart>
</c:chartSpace>)";


constexpr std::string_view kStyles = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>)";

constexpr std::string_view kDocument = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <w:body>

    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">Report: </w:t></w:r>
      <w:bookmarkStart w:id="1" w:name="_Header.ReportName"/>
      <w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t>{{title}}</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>

    <w:p>
      <w:r><w:t xml:space="preserve">Prepared by: </w:t></w:r>
      <w:bookmarkStart w:id="2" w:name="_Header.User"/>
      <w:r><w:t>{{user}}</w:t></w:r>
      <w:bookmarkEnd w:id="2"/>
    </w:p>

    <w:p>
      <w:r><w:t xml:space="preserve">Date: </w:t></w:r>
      <w:bookmarkStart w:id="3" w:name="_Header.UKDate"/>
      <w:r><w:t>{{date}}</w:t></w:r>
      <w:bookmarkEnd w:id="3"/>
    </w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Greeting</w:t></w:r>
    </w:p>

    <w:p>
      <w:bookmarkStart w:id="4" w:name="Body.Greeting"/>
      <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
      <w:r><w:t>{{name}}</w:t></w:r>
      <w:r><w:t xml:space="preserve">, welcome to </w:t></w:r>
      <w:r><w:t>{{app}}</w:t></w:r>
      <w:r><w:t>.</w:t></w:r>
      <w:bookmarkEnd w:id="4"/>
    </w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Logo</w:t></w:r>
    </w:p>

    <w:p>
      <w:bookmarkStart w:id="6" w:name="Body.Logo"/>
      <w:r><w:t>{{logo}}</w:t></w:r>
      <w:bookmarkEnd w:id="6"/>
    </w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Scores</w:t></w:r>
    </w:p>

    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top     w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left    w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:bottom  w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:right   w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="4000"/>
        <w:gridCol w:w="2000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Name</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Score</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:bookmarkStart w:id="5" w:name="Body.Scores"/>
            <w:r><w:t>{{rowName}}</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:r><w:t>{{rowScore}}</w:t></w:r>
            <w:bookmarkEnd w:id="5"/>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Stats</w:t></w:r>
    </w:p>

    <w:p><w:r><w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="5030470" cy="3200400"/>
        <wp:docPr id="1" name="Chart 1"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rIdChart1"/>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>

    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Browser Share</w:t></w:r>
    </w:p>

    <w:p><w:r><w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="4114800" cy="3200400"/>
        <wp:docPr id="2" name="Chart 2"/>
        <wp:cNvGraphicFramePr/>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rIdChart2"/>
        </a:graphicData></a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>)";

fs::path make_template(const fs::path& out) {
    const std::vector<Part> parts = {
        {"[Content_Types].xml",          std::string(kContentTypes)},
        {"_rels/.rels",                  std::string(kPackageRels)},
        {"word/_rels/document.xml.rels", std::string(kDocumentRels)},
        {"word/styles.xml",              std::string(kStyles)},
        {"word/document.xml",            std::string(kDocument)},
        {"word/charts/chart1.xml",       std::string(kBarChart)},
        {"word/charts/chart2.xml",       std::string(kPieChart)},
    };
    if (write_docx(out, parts) != 0) {
        throw std::runtime_error("failed to write template");
    }
    return out;
}

// Force the Microsoft Word and LibreOffice branches off regardless of what
// happens to be installed on this machine — the whole point of this example
// is showing the native PoDoFo path runs without either converter. No effect
// on the process outside this run.
void disable_external_converters() {
#if defined(_WIN32)
    _putenv_s("DOCWEFT_NO_MSWORD", "1");
    _putenv_s("DOCWEFT_SOFFICE", "");
#else
    setenv("DOCWEFT_NO_MSWORD", "1", 1);
    setenv("DOCWEFT_SOFFICE", "", 1);
#endif
}

} // namespace

int main(int argc, char** argv) {
    const fs::path out_pdf = (argc > 1) ? argv[1] : "native_report.pdf";
    const fs::path template_path = "native_template.docx";
    const fs::path logo_path = "logo.png";

    disable_external_converters();

    try {
        make_template(template_path);

        auto m = docweft::make_docx_merger();
        m->setCodePage("UTF-8");
        m->load(template_path.string());

        m->setClipboardValue("_Header.ReportName", "{{title}}", "Q1 2026 Summary");
        m->setClipboardValue("_Header.User",       "{{user}}",  "Иван Петров");
        m->setClipboardValue("_Header.UKDate",     "{{date}}",  "2026-04-22");
        m->setClipboardValue("Body.Greeting", "{{name}}", "Alice");
        m->setClipboardValue("Body.Greeting", "{{app}}",  "Acme Corp");
        m->setTableRow("Body.Scores",
            {"{{rowName}}", "{{rowScore}}"},
            {
                {"Alice",    "95"},
                {"Борис",    "82"},
                {"Carol Жу", "77"},
            });
        m->setImage("Body.Logo", logo_path,
                    docweft::ImageSize{/*min*/64, 64, /*max*/256, 256});

        m->save(out_pdf.string());
        std::cout << "Wrote " << out_pdf
                  << " via the native PoDoFo backend — no Microsoft Word or "
                     "LibreOffice were involved.\n";
    } catch (const docweft::ReportException& e) {
        std::cerr << "[DocWeft error] " << docweft::to_string(e.code())
                   << ": " << e.what() << "\n";
        if (e.code() == docweft::ReportError::NoConverter) {
            std::cerr << "This binary needs the library built with "
                         "-DDOCWEFT_ENABLE_NATIVE_PDF=ON (requires PoDoFo) "
                         "— see PLAN.md.\n";
        }
        return 1;
    }
    return 0;
}
