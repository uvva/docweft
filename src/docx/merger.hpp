#pragma once

#include "docweft/merger.hpp"

#include <pugixml.hpp>

#include <cstdint>
#include <filesystem>
#include <string>
#include <unordered_map>
#include <vector>

namespace docweft::docx {

/// Name of the main document part inside every .docx archive.
inline constexpr const char* kDocumentPart = "word/document.xml";

/// In-memory representation of a loaded .docx.
/// We keep the full archive contents as a (path → bytes) map so we can
/// rewrite any part and repack on save() while preserving everything else
/// (styles, fonts, images, chart xml, etc.).
class DocxMerger final : public IReportMerger {
public:
    DocxMerger();
    ~DocxMerger() override;

    // ── IReportMerger ────────────────────────────────────────────────────
    void setCodePage(const std::string& code_page) override;
    void load(const std::string& path) override;
    void save(const std::string& path) override;

    void setClipboardValue(const std::string& bookmark,
                           const std::string& field,
                           const std::string& value) override;

    [[nodiscard]] bool hasBookmark(const std::string& bookmark) const override;
    void clearBookmark(const std::string& bookmark) override;

    void setTableRow(const std::string& bookmark,
                     const std::vector<std::string>& fields,
                     const std::vector<std::vector<std::string>>& rows) override;

    void setImage(const std::string& bookmark,
                  const std::filesystem::path& image_path,
                  const ImageSize& bounds = {}) override;

    void setChartValue(const std::string& bookmark,
                       const std::string& field,
                       const std::string& series,
                       const std::string& category,
                       double             value) override;

    void setChartSeriesName(const std::string& bookmark,
                            const std::string& old_name,
                            const std::string& new_name) override;

    void setChartTitle(const std::string& bookmark,
                       const std::string& title) override;

    void setChartAxisTitle(const std::string& bookmark,
                           ChartAxis           axis,
                           const std::string&  title) override;

    void setChartData(const std::string&             bookmark,
                      const std::vector<std::string>& categories,
                      const std::vector<ChartSeries>& series) override;

    void paste(const std::string& bookmark) override;

    // ── Introspection (used by tests) ────────────────────────────────────
    [[nodiscard]] bool loaded() const noexcept { return loaded_; }
    [[nodiscard]] std::string document_xml() const;
    [[nodiscard]] const std::vector<std::string>& pasted_bookmarks() const noexcept {
        return pasted_;
    }

private:
    /// Raw archive: part name (zip path) → file bytes.
    std::unordered_map<std::string, std::string> parts_;

    /// Parsed word/document.xml (re-serialized into parts_ on save()).
    pugi::xml_document document_;

    bool        loaded_     = false;
    std::string code_page_  = "UTF-8";
    std::vector<std::string> pasted_;

    // ── Internals ────────────────────────────────────────────────────────
    void read_archive(const std::filesystem::path& path);
    void write_archive(const std::filesystem::path& path) const;
    static void write_archive(
        const std::unordered_map<std::string, std::string>& parts,
        const std::filesystem::path&                         path);
    void reparse_document();
    void reserialize_document();

    /// Find all <w:bookmarkStart> nodes matching `name`.
    /// Returns nodes in document order.
    std::vector<pugi::xml_node> find_bookmark_starts(const std::string& name) const;

    /// Resolve `bookmark`'s enclosing <w:p>'s embedded <c:chart> down to
    /// the archive part name holding its chart XML (e.g.
    /// "word/charts/chart1.xml"). Shared entry point for every
    /// setChart*() method.
    /// Throws ReportException{InvalidBookmark | CantCopyDocxTemplate}.
    [[nodiscard]] std::string locate_chart_part(const std::string& bookmark) const;
};

} // namespace docweft::docx
