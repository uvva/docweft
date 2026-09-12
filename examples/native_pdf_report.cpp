// native_pdf_report — same kind of report as basic_report, but demonstrates
// TextFabric's native PoDoFo PDF backend (TEXTFABRIC_ENABLE_NATIVE_PDF): the
// resulting PDF is produced with no Microsoft Word or LibreOffice involved
// at all, not even installed on the machine. See PLAN.md for the backend's
// scope and known gaps (most notably: no embedded charts — the template
// built here deliberately has none, unlike generate_template's).
//
// This binary only exists when the library was built with
// -DTEXTFABRIC_ENABLE_NATIVE_PDF=ON (see examples/CMakeLists.txt) — save()
// would otherwise throw NoConverter once Word/LibreOffice are disabled
// below, which is the whole point of this example.
//
// Usage:
//   ./native_pdf_report [output.pdf]    # default: native_report.pdf

#include <textfabric/error.hpp>
#include <textfabric/merger.hpp>

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

// Same boilerplate parts as generate_template.cpp, minus everything chart-
// related (no [Content_Types].xml chart override, no chart relationship, no
// word/charts/chart1.xml, no Body.Stats paragraph) — the native PDF backend
// doesn't render charts, so this template deliberately has none.

constexpr std::string_view kContentTypes = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
</Relationships>)";

constexpr std::string_view kStyles = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>)";

constexpr std::string_view kDocument = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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

    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>)";

fs::path make_chartless_template(const fs::path& out) {
    const std::vector<Part> parts = {
        {"[Content_Types].xml",          std::string(kContentTypes)},
        {"_rels/.rels",                  std::string(kPackageRels)},
        {"word/_rels/document.xml.rels", std::string(kDocumentRels)},
        {"word/styles.xml",              std::string(kStyles)},
        {"word/document.xml",            std::string(kDocument)},
    };
    if (write_docx(out, parts) != 0) {
        throw std::runtime_error("failed to write template");
    }
    return out;
}

// Force the LibreOffice branch off regardless of what happens to be
// installed on this machine — the whole point of this example is showing
// the native PoDoFo path runs without either converter. No effect on
// Microsoft Word detection (Windows-only, already skipped here since we
// don't touch the registry probe), and no effect on the process outside
// this run.
void disable_external_converters() {
#if defined(_WIN32)
    _putenv_s("TEXTFABRIC_SOFFICE", "");
#else
    setenv("TEXTFABRIC_SOFFICE", "", 1);
#endif
}

} // namespace

int main(int argc, char** argv) {
    const fs::path out_pdf = (argc > 1) ? argv[1] : "native_report.pdf";
    const fs::path template_path = "native_template.docx";
    const fs::path logo_path = "logo.png";

    disable_external_converters();

    try {
        make_chartless_template(template_path);

        auto m = textfabric::make_docx_merger();
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
                    textfabric::ImageSize{/*min*/64, 64, /*max*/256, 256});

        m->save(out_pdf.string());
        std::cout << "Wrote " << out_pdf
                  << " via the native PoDoFo backend — no Microsoft Word or "
                     "LibreOffice were involved.\n";
    } catch (const textfabric::ReportException& e) {
        std::cerr << "[TextFabric error] " << textfabric::to_string(e.code())
                   << ": " << e.what() << "\n";
        if (e.code() == textfabric::ReportError::NoConverter) {
            std::cerr << "This binary needs the library built with "
                         "-DTEXTFABRIC_ENABLE_NATIVE_PDF=ON (requires PoDoFo) "
                         "— see PLAN.md.\n";
        }
        return 1;
    }
    return 0;
}
