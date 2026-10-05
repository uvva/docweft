// generate_template — produces a minimal Word-compatible .docx template
// exercising every bookmark kind the library supports:
//
//   _Header.User        — contains "{{user}}"
//   _Header.ReportName  — contains "{{title}}"
//   _Header.UKDate      — contains "{{date}}"
//   Body.Greeting       — contains "Hello {{name}}, welcome to {{app}}."
//   Body.Scores         — inside a <w:tr> template row with two columns:
//                         "{{rowName}}" and "{{rowScore}}" — cloned per
//                         data row by IReportMerger::setTableRow.
//   Body.Logo           — holds "{{logo}}" as a placeholder inside a
//                         paragraph; IReportMerger::setImage replaces it
//                         with the provided PNG.
//   Body.Stats          — wraps a <w:drawing> that references an embedded
//                         bar chart (word/charts/chart1.xml). Initial series
//                         "Measurements" × {"Region A/B/C"} holds zeros so
//                         the example can overwrite them via setChartValue.
//
// The result is a valid OOXML zip that opens in MS Word / LibreOffice.
// Usage:
//   ./generate_template [output.docx]    # default: template.docx

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
        if (!src) {
            std::cerr << "zip_source_buffer failed for " << p.name << "\n";
            zip_discard(z);
            return 1;
        }
        if (zip_file_add(z, p.name.c_str(), src,
                          ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) < 0) {
            std::cerr << "zip_file_add failed for " << p.name << ": "
                      << zip_strerror(z) << "\n";
            zip_source_free(src);
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

// ── OOXML parts ────────────────────────────────────────────────────────────

constexpr std::string_view kContentTypes = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>)";

constexpr std::string_view kPackageRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="word/document.xml"/>
</Relationships>)";

// Document settings as a new Word 2013+ document has them. Without this
// part a converter falls back to its own defaults — LibreOffice then adds
// up the space after one paragraph and before the next, which Word (and
// LibreOffice for a document that has settings) doesn't.
constexpr std::string_view kSettings = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
</w:settings>)";

constexpr std::string_view kDocumentRels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rIdSettings"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
                Target="settings.xml"/>
  <Relationship Id="rIdChart1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                Target="charts/chart1.xml"/>
</Relationships>)";

// Bar chart: one series "Measurements" × three categories ("Region A/B/C").
// Initial values are zeros on purpose — the example overwrites them at
// runtime via setChartValue, and non-zero numbers in the rendered chart
// prove end-to-end correctness of Stage 4.
constexpr std::string_view kChart = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache>
            <c:ptCount val="1"/>
            <c:pt idx="0"><c:v>Measurements</c:v></c:pt>
          </c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache>
            <c:ptCount val="3"/>
            <c:pt idx="0"><c:v>Region A</c:v></c:pt>
            <c:pt idx="1"><c:v>Region B</c:v></c:pt>
            <c:pt idx="2"><c:v>Region C</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache>
            <c:formatCode>General</c:formatCode>
            <c:ptCount val="3"/>
            <c:pt idx="0"><c:v>0</c:v></c:pt>
            <c:pt idx="1"><c:v>0</c:v></c:pt>
            <c:pt idx="2"><c:v>0</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111111111"/>
        <c:axId val="222222222"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="111111111"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="222222222"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
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

    <w:p>
      <w:bookmarkStart w:id="7" w:name="Body.Stats"/>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="5486400" cy="3200400"/>
            <wp:docPr id="1" name="Chart 1"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rIdChart1"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
      <w:bookmarkEnd w:id="7"/>
    </w:p>

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>)";

} // namespace

int main(int argc, char** argv) {
    const fs::path out = (argc > 1) ? argv[1] : "template.docx";

    std::vector<Part> parts = {
        {"[Content_Types].xml",          std::string(kContentTypes)},
        {"_rels/.rels",                  std::string(kPackageRels)},
        {"word/_rels/document.xml.rels", std::string(kDocumentRels)},
        {"word/styles.xml",              std::string(kStyles)},
        {"word/settings.xml",            std::string(kSettings)},
        {"word/document.xml",            std::string(kDocument)},
        {"word/charts/chart1.xml",       std::string(kChart)},
    };

    if (int rc = write_docx(out, parts); rc != 0) return rc;

    std::cout << "Wrote template: " << out << "\n"
              << "  bookmarks: _Header.ReportName, _Header.User, _Header.UKDate,\n"
              << "             Body.Greeting, Body.Scores (table row), Body.Logo,\n"
              << "             Body.Stats (bar chart)\n"
              << "  fields:    {{title}}, {{user}}, {{date}}, {{name}}, {{app}},\n"
              << "             {{rowName}}, {{rowScore}}, {{logo}}\n"
              << "  chart:     series 'Measurements' × {Region A, Region B, Region C}\n"
              << "             (all values zero in the template — overwrite via setChartValue)\n";
    return 0;
}
