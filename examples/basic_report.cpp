// basic_report — the canonical consumer-side example of DocWeft.
//
// Loads a template .docx, fills four header bookmarks with runtime values,
// clones a table row three times with per-row data, and saves the result
// as .docx. If LibreOffice (soffice) is available, additionally writes
// .pdf and .html alongside.
//
// Usage:
//   ./basic_report [template.docx] [output.docx]
//
// Defaults: template.docx → report.docx (.pdf / .html are derived from the
// same stem and skipped gracefully if soffice is absent).

#include <docweft/error.hpp>
#include <docweft/merger.hpp>
#include <docweft/docweft.hpp>   // version

#include <array>
#include <chrono>
#include <ctime>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

int main(int argc, char** argv) {
    const std::string tpl_path = (argc > 1) ? argv[1] : "template.docx";
    const std::string out_path = (argc > 2) ? argv[2] : "report.docx";

    std::cout << "DocWeft v" << docweft::version << "\n"
              << "  template: " << tpl_path << "\n"
              << "  output:   " << out_path << "\n\n";

    // ── 1. Build the merger via the library factory ──────────────────────
    auto merger = docweft::make_docx_merger();

    try {
        merger->setCodePage("UTF-8");

        // ── 2. Load the template ─────────────────────────────────────────
        merger->load(tpl_path);

        // ── 3. Substitute fields inside bookmarks ────────────────────────
        merger->setClipboardValue("_Header.ReportName", "{{title}}", "Q1 2026 Summary");
        merger->setClipboardValue("_Header.User",       "{{user}}",  "Иван Петров");

        // Timestamp the report so each run produces a new document.
        const auto now_t = std::chrono::system_clock::to_time_t(
                               std::chrono::system_clock::now());
        std::array<char, 16> buf{};
        std::strftime(buf.data(), buf.size(), "%Y-%m-%d", std::localtime(&now_t));
        const std::string date(buf.data());
        merger->setClipboardValue("_Header.UKDate", "{{date}}", date);

        // Multi-field bookmark — both placeholders sit inside the same block.
        merger->setClipboardValue("Body.Greeting", "{{name}}", "Alice");
        merger->setClipboardValue("Body.Greeting", "{{app}}",  "Acme Corp");

        // ── 4a. Embed the sample logo from examples/logo.png (Stage 5). ──
        // The file ships with the repo and is copied next to basic_report by
        // CMake (see examples/CMakeLists.txt). Resolve it relative to the
        // output path so relative-CWD and absolute-path invocations both work.
        namespace fs = std::filesystem;
        const fs::path logo_path = fs::path(out_path).replace_filename("logo.png");
        const bool logo_embedded = fs::exists(logo_path);
        if (!logo_embedded) {
            std::cerr << "[DocWeft] warning: " << logo_path.string()
                      << " not found — skipping setImage demo. "
                         "Run this binary from examples/ in the build tree, "
                         "or copy examples/logo.png next to it.\n";
        } else {
            // Clamp the displayed size: the sample PNG is only 2×2 pixels,
            // which would render as a dust speck in Word. min_width_px=64
            // scales it up uniformly to 64×64, preserving the 1:1 aspect.
            merger->setImage("Body.Logo", logo_path,
                             docweft::ImageSize{/*min_w*/64, /*min_h*/64,
                                                   /*max_w*/256, /*max_h*/256});
        }

        // ── 4. Table rows (Stage 3 — clone <w:tr> per data entry) ────────
        // The template has one <w:tr> marked with the `Body.Scores`
        // bookmark and two placeholders ({{rowName}}, {{rowScore}}).
        // setTableRow clones it once per entry, replaces the placeholders
        // in the matching cells, and drops the template row itself.
        const std::vector<std::vector<std::string>> scores = {
            {"Alice",    "95"},
            {"Борис",    "82"},
            {"Carol Жу", "77"},
        };
        merger->setTableRow("Body.Scores",
                            {"{{rowName}}", "{{rowScore}}"},
                            scores);

        // ── 4b. Populate the embedded bar chart (Stage 4). ───────────────
        // The template ships `Body.Stats` with series "Measurements" and
        // zeros for Region A/B/C — overwriting them here proves that the
        // rendered PDF/HTML shows real numbers, i.e. Word / LibreOffice
        // actually reads the <c:numCache> we rewrite.
        const std::vector<std::pair<std::string, double>> stats = {
            {"Region A", 42.0},
            {"Region B", 73.5},
            {"Region C", 55.0},
        };
        for (const auto& [region, value] : stats) {
            merger->setChartValue("Body.Stats", "ChartField",
                                  "Measurements", region, value);
        }

        // ── 5. Activate sections (v1 = no-op + validation) ───────────────
        merger->paste("_Header.ReportName");
        merger->paste("_Header.User");
        merger->paste("_Header.UKDate");
        merger->paste("Body.Greeting");
        merger->paste("Body.Logo");
        merger->paste("Body.Stats");

        // ── 6. Save primary .docx ────────────────────────────────────────
        merger->save(out_path);

        std::cout << "Wrote " << out_path << "\n"
                  << "Open it in MS Word / LibreOffice to verify:\n"
                  << "  - title 'Q1 2026 Summary'\n"
                  << "  - author 'Иван Петров'\n"
                  << "  - today's date (" << date << ")\n"
                  << "  - greeting 'Hello Alice, welcome to Acme Corp.'\n"
                  << "  - Scores table with " << scores.size() << " rows:\n";
        for (const auto& row : scores) {
            std::cout << "      " << row[0] << " — " << row[1] << "\n";
        }
        if (logo_embedded) {
            std::cout << "  - embedded PNG logo from " << logo_path.string()
                      << " (via setImage, upscaled to 64×64 via "
                         "ImageSize{min=64, max=256})\n";
        } else {
            std::cout << "  - logo placeholder {{logo}} left intact "
                         "(logo.png was missing next to the exe)\n";
        }
        std::cout << "  - bar chart 'Measurements' with "
                  << stats.size() << " categories:\n";
        for (const auto& [region, value] : stats) {
            std::cout << "      " << region << " — " << value << "\n";
        }

        // ── 7. Optional PDF / HTML exports via LibreOffice (Stage 6) ─────
        // If soffice isn't installed we get ReportError::NoConverter and
        // print a hint instead of failing — this keeps the example usable
        // on machines without LibreOffice.
        const fs::path out_fs(out_path);
        const fs::path pdf_path  = fs::path(out_fs).replace_extension(".pdf");
        const fs::path html_path = fs::path(out_fs).replace_extension(".html");

        for (const auto& [target, label] :
             std::initializer_list<std::pair<fs::path, const char*>>{
                 {pdf_path,  "PDF"},
                 {html_path, "HTML"}}) {
            try {
                merger->save(target.string());
                std::cout << "Wrote " << target.string()
                          << " (" << label
                          << " via Word or LibreOffice, whichever was found)\n";
            } catch (const docweft::ReportException& e) {
                if (e.code() == docweft::ReportError::NoConverter) {
                    std::cout << label
                              << ": skipped — no converter available. "
                                 "Install Microsoft Word (Windows) or "
                                 "LibreOffice to enable PDF/HTML export.\n";
                } else {
                    throw;  // real failure — bubble out
                }
            }
        }
        return 0;
    }
    catch (const docweft::ReportException& e) {
        std::cerr << "[DocWeft error] "
                  << docweft::to_string(e.code()) << ": "
                  << e.what() << "\n";
        return 1;
    }
    catch (const std::exception& e) {
        std::cerr << "[std::exception] " << e.what() << "\n";
        return 2;
    }
}
