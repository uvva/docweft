#pragma once

#include "textfabric/export.h"

#include <cstdint>
#include <filesystem>
#include <memory>
#include <string>
#include <vector>

namespace textfabric {

/// Display-size bounds for `setImage`. All four fields are in pixels at a
/// 96 DPI baseline (the same unit Word uses when importing images through
/// drag-and-drop). A value of `0` means "no constraint on this dimension".
/// Aspect ratio is always preserved — the merger picks a single uniform
/// scale that satisfies every non-zero bound simultaneously.
///
/// Resolution rules (see `IReportMerger::setImage` for the full algorithm):
///   - all zeros    → image is inserted at its intrinsic pixel size;
///   - image fits   → intrinsic size is kept;
///   - image bigger → shrunk to the tightest of the max bounds;
///   - image smaller → enlarged to the loosest of the min bounds;
///   - min > max    → throws `ReportError::InvalidField` (conflicting bounds).
struct ImageSize {
    std::uint32_t min_width_px  = 0;
    std::uint32_t min_height_px = 0;
    std::uint32_t max_width_px  = 0;
    std::uint32_t max_height_px = 0;
};

/// Abstract merger API for filling DOCX templates by bookmark name.
///
/// Implementations:
///   - docx::DocxMerger — uses libzip + pugixml directly.
///
/// All string parameters are UTF-8.
class TEXTFABRIC_API IReportMerger {
public:
    virtual ~IReportMerger() = default;

    /// Set the code page used for input/output. Only "UTF-8" is
    /// supported in v1 — anything else throws ReportException.
    virtual void setCodePage(const std::string& code_page) = 0;

    /// Load a DOCX template from disk.
    /// Throws ReportException{CantOpenTemplate | CantCopyDocxTemplate}.
    virtual void load(const std::string& path) = 0;

    /// Save the current document to disk.
    /// Format is inferred from the file extension:
    ///   .docx → native save
    ///   .pdf / .html → first working converter of: Microsoft Word (Windows,
    ///                  macOS), LibreOffice, native PDF renderer
    ///                  (TEXTFABRIC_ENABLE_NATIVE_PDF, .pdf only), remote
    ///                  converter (TEXTFABRIC_ENABLE_REMOTE_CONVERTER +
    ///                  TEXTFABRIC_CONVERTER_URL, .pdf only)
    /// Throws ReportException{SaveFailed | NoConverter | NotImplemented}.
    virtual void save(const std::string& path) = 0;

    /// Substitute a text field inside a bookmark.
    /// `bookmark` — DOCX bookmark name (e.g. "_Header.User")
    /// `field`    — placeholder token inside the bookmark (e.g. "{{user}}" or "kUSER")
    /// `value`    — replacement string (UTF-8)
    /// Throws ReportException{InvalidBookmark | InvalidField}.
    virtual void setClipboardValue(const std::string& bookmark,
                                   const std::string& field,
                                   const std::string& value) = 0;

    /// Query whether a bookmark with this exact name exists in the loaded
    /// template. Returns false both when the template has none and when no
    /// template is loaded yet — never throws. Use it to validate a
    /// template's bookmark set up front, or to build a diagnostic list of
    /// bookmarks a report run never touched.
    ///
    /// Does not resolve "_Header"/"_Footer" pseudo-bookmarks (see
    /// setClipboardValue above) — those aren't backed by a real
    /// `<w:bookmarkStart>` to look up.
    [[nodiscard]] virtual bool hasBookmark(const std::string& bookmark) const = 0;

    /// Erase the placeholder text inside a bookmark, leaving it empty.
    ///
    /// Use this for a bookmark the caller has decided not to fill on this
    /// run (e.g. an optional section with no data for it) — without an
    /// explicit call to either this or setClipboardValue, the original
    /// template placeholder text (including literal "{field}" tokens)
    /// survives untouched into the saved document.
    ///
    /// Does not resolve "_Header"/"_Footer" pseudo-bookmarks — see
    /// setClipboardValue above.
    /// Throws ReportException{InvalidBookmark}.
    virtual void clearBookmark(const std::string& bookmark) = 0;

    /// Clone a table row and fill its cells.
    ///
    /// The bookmark must live inside a `<w:tr>` template row that serves as
    /// the prototype. For each element of `rows`, the row is cloned and the
    /// `fields[j]` placeholder in column j is replaced by `rows[i][j]`.
    /// The template row itself is removed after cloning; an empty `rows`
    /// therefore yields an empty body for that table. Bookmarks inside
    /// cloned rows are stripped to avoid duplicate `w:id` attributes.
    ///
    /// `bookmark` — DOCX bookmark inside the template `<w:tr>`.
    /// `fields`   — column placeholders (e.g. `{"{{name}}", "{{score}}"}`).
    /// `rows`     — values; `rows[i].size()` must equal `fields.size()`.
    /// Throws ReportException{InvalidBookmark | InvalidField}.
    virtual void setTableRow(const std::string& bookmark,
                             const std::vector<std::string>& fields,
                             const std::vector<std::vector<std::string>>& rows) = 0;

    /// Embed a raster image at `bookmark`.
    ///
    /// Accepted formats are detected by magic bytes:
    ///   - PNG: passes through as-is and goes into `word/media/image*.png`.
    ///   - JPEG / BMP / TIFF: transcoded to PNG at call time
    ///     (requires the stb_image + libtiff support enabled in the build).
    ///
    /// The merger inserts a `<w:drawing>` element inside the paragraph that
    /// holds `bookmark`, displacing any text the bookmark already contained
    /// (so the template can carry a placeholder like "{{image}}" that gets
    /// fully replaced by the picture). A new `<Relationship Id="rIdX" ... />`
    /// is added to `word/_rels/document.xml.rels` and `[Content_Types].xml`
    /// gains a `<Default Extension="png" ContentType="image/png"/>` entry
    /// if not already present.
    ///
    /// Image dimensions are read from the source file; the displayed size in
    /// `<wp:extent>` defaults to 1 px ≈ 9525 EMU (96 DPI baseline). The
    /// optional `bounds` argument clamps the displayed size between user-
    /// supplied minima/maxima while preserving aspect ratio — see `ImageSize`.
    ///
    /// Throws `ReportException`:
    ///   - InvalidBookmark:        bookmark not found or not inside `<w:p>`
    ///   - InvalidField:           `bounds` are internally contradictory
    ///                             (e.g. min_width > max_width, or the
    ///                             aspect-preserving scale range is empty)
    ///   - CantOpenTemplate:       image file missing / unreadable
    ///   - CantCopyDocxTemplate:   file is not a recognised image format
    ///   - NotImplemented:         JPEG/BMP/TIFF hit the baseline (PNG-only)
    ///                             build — convert to PNG first or enable the
    ///                             Phase 2 decoders.
    virtual void setImage(const std::string& bookmark,
                          const std::filesystem::path& image_path,
                          const ImageSize& bounds = {}) = 0;

    /// Overwrite the value at an existing (series, category) point inside
    /// an embedded chart anchored at `bookmark`. `field` is accepted for
    /// API compatibility but not used for matching.
    ///
    /// This only rewrites a point that's already in the chart's data
    /// cache — it cannot add or remove a category or series. The chart's
    /// category axis and series set are whatever the template's author
    /// laid out in Excel; if the number of categories is only known at
    /// runtime and doesn't match the template, calls for the missing
    /// categories throw InvalidField. There's no API here to replace the
    /// axis/series shape itself — only to update values already on it.
    ///
    /// Throws ReportException:
    ///   - InvalidBookmark: bookmark not found, not in a <w:p>, or that
    ///                      paragraph has no chart
    ///   - InvalidField:    `series` or `category` don't exist in the chart
    ///   - NotImplemented:  chart is scatter/bubble/stock/surface — only
    ///                      cat/val shapes (bar/line/pie/area/radar/
    ///                      doughnut/3D variants) are supported
    virtual void setChartValue(const std::string& bookmark,
                               const std::string& field,
                               const std::string& series,
                               const std::string& category,
                               double             value) = 0;

    /// Rename an existing chart series in place — same (series, category)
    /// data points, new display name. Does not add or remove a series;
    /// use setChartData for that.
    ///
    /// Throws ReportException:
    ///   - InvalidBookmark: bookmark not found, not in a <w:p>, or that
    ///                      paragraph has no chart
    ///   - InvalidField:    `old_name` doesn't match any series in the chart
    ///   - NotImplemented:  chart is scatter/bubble/stock/surface
    virtual void setChartSeriesName(const std::string& bookmark,
                                    const std::string& old_name,
                                    const std::string& new_name) = 0;

    /// Which axis setChartAxisTitle targets.
    enum class ChartAxis { Category, Value };

    /// Set the chart's title text, replacing whatever the template author
    /// set in Excel/Word (including "no title" — autoTitleDeleted is
    /// cleared and a title element is created if the chart didn't have
    /// one). If the chart already had a title, its box formatting, first
    /// run's font/color, and paragraph alignment carry over onto the new
    /// text — only the string itself is replaced, not the styling. (A
    /// title with several differently-formatted runs still collapses to
    /// one run using the first run's style, since this call sets one
    /// plain string.)
    ///
    /// Throws ReportException:
    ///   - InvalidBookmark: bookmark not found, not in a <w:p>, or that
    ///                      paragraph has no chart
    ///   - NotImplemented:  chart is scatter/bubble/stock/surface
    virtual void setChartTitle(const std::string& bookmark,
                               const std::string& title) = 0;

    /// Set the title text of the chart's category or value axis.
    /// Same title-creation and formatting-preservation behavior as
    /// setChartTitle when the axis already has a title, or has none yet.
    ///
    /// Throws ReportException:
    ///   - InvalidBookmark: bookmark not found, not in a <w:p>, or that
    ///                      paragraph has no chart
    ///   - InvalidField:    chart has no axis of the requested kind (e.g.
    ///                      a pie chart has neither)
    ///   - NotImplemented:  chart is scatter/bubble/stock/surface
    virtual void setChartAxisTitle(const std::string& bookmark,
                                   ChartAxis           axis,
                                   const std::string&  title) = 0;

    /// One named series' values for setChartData, in the same order as
    /// the `categories` vector passed alongside it.
    struct ChartSeries {
        std::string         name;
        std::vector<double> values;
    };

    /// Replace a chart's entire category axis and series set — unlike
    /// setChartValue, this can add or remove categories/series, not just
    /// overwrite existing points.
    ///
    /// `categories` becomes the new category axis, in order. Each entry of
    /// `series` supplies a name and one value per category (`values.size()`
    /// must equal `categories.size()`). The template's existing series
    /// supply the visual style (color/marker/line) — the first `series.size()`
    /// of them are reused in template order and restyled with the new
    /// name/data; if `series` has more entries than the template had series,
    /// the extra ones clone the last template series' style; if fewer, the
    /// surplus template series are removed. Per-point color overrides the
    /// template had on individual categories don't survive — they're keyed
    /// to point indices from the old category axis, which have no
    /// principled mapping onto a reshaped one, so every point falls back
    /// to the chart's normal per-index theme-color rotation.
    ///
    /// When the chart carries an embedded workbook (word/embeddings/*.xlsx,
    /// linked via <c:externalData>), its backing worksheet is rewritten to
    /// match — so "Edit Data in Excel"/"Refresh Data" in Word stays
    /// consistent with what's on screen. A chart with no embedded workbook
    /// only gets its cache rewritten, same as setChartValue.
    ///
    /// Throws ReportException:
    ///   - InvalidBookmark: bookmark not found, not in a <w:p>, or that
    ///                      paragraph has no chart
    ///   - InvalidField:    `categories` or some `series[i].values` is
    ///                      empty, or a values vector's length doesn't
    ///                      match `categories.size()`
    ///   - CantCopyDocxTemplate: template's chart has zero template series
    ///                      to clone style from, or its embedded workbook
    ///                      is present but not a well-formed .xlsx
    ///   - NotImplemented:  chart is scatter/bubble/stock/surface
    virtual void setChartData(const std::string&              bookmark,
                              const std::vector<std::string>&  categories,
                              const std::vector<ChartSeries>&  series) = 0;

    /// Paste/activate a template section anchored at a bookmark.
    /// Calling paste() materializes the currently-staged field values
    /// into the document.
    /// Throws ReportException{InvalidBookmark}.
    virtual void paste(const std::string& bookmark) = 0;
};

/// Factory: builds a DOCX-backed merger.
[[nodiscard]] TEXTFABRIC_API std::unique_ptr<IReportMerger> make_docx_merger();

} // namespace textfabric
