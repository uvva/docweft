#include "docx/pdf_native.hpp"
#include "textfabric/error.hpp"

#include <fmt/format.h>

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#if defined(TEXTFABRIC_HAVE_PODOFO)
#include <podofo/podofo.h>
#endif

namespace textfabric::docx {

#if defined(TEXTFABRIC_HAVE_PODOFO)

namespace {

using RelMap  = std::unordered_map<std::string, std::string>;
using PartMap = std::unordered_map<std::string, std::string>;

double twips_to_pt(long long twips) { return static_cast<double>(twips) / 20.0; }
double emu_to_pt(double emu)        { return emu / 12700.0; }

struct PageMetrics {
    double width_pt        = 595.28;  // A4 default, overridden by <w:pgSz> below
    double height_pt       = 841.89;
    double margin_left_pt  = 72.0;    // 1", overridden by <w:pgMar> below
    double margin_right_pt = 72.0;
    double margin_top_pt   = 72.0;
    double margin_bottom_pt = 72.0;
};

PageMetrics read_page_metrics(pugi::xml_node body) {
    PageMetrics pm;
    pugi::xml_node sect = body.child("w:sectPr");
    if (pugi::xml_node sz = sect.child("w:pgSz")) {
        pm.width_pt  = twips_to_pt(sz.attribute("w:w").as_llong(11906));
        pm.height_pt = twips_to_pt(sz.attribute("w:h").as_llong(16838));
    }
    if (pugi::xml_node mar = sect.child("w:pgMar")) {
        pm.margin_left_pt   = twips_to_pt(mar.attribute("w:left").as_llong(1440));
        pm.margin_right_pt  = twips_to_pt(mar.attribute("w:right").as_llong(1440));
        pm.margin_top_pt    = twips_to_pt(mar.attribute("w:top").as_llong(1440));
        pm.margin_bottom_pt = twips_to_pt(mar.attribute("w:bottom").as_llong(1440));
    }
    return pm;
}

// pugixml is namespace-unaware, so descendant search compares the literal
// prefixed tag name — same pattern as find_descendant_chart() in merger.cpp.
pugi::xml_node find_descendant(pugi::xml_node root, const char* tag) {
    struct Finder : pugi::xml_tree_walker {
        const char* tag;
        pugi::xml_node hit;
        explicit Finder(const char* t) : tag(t) {}
        bool for_each(pugi::xml_node& n) override {
            if (std::strcmp(n.name(), tag) == 0) {
                hit = n;
                return false;
            }
            return true;
        }
    } finder(tag);
    root.traverse(finder);
    return finder.hit;
}

struct RunStyle {
    bool   bold    = false;
    double size_pt = 11.0;
};

// Only the first run's formatting is consulted — same "first run wins"
// convention the rest of the merger already uses for setClipboardValue.
// A paragraph that changes style mid-run (e.g. one bold word in a plain
// sentence) renders uniformly in the first run's style; this is a known,
// documented simplification (see PLAN.md).
RunStyle first_run_style(pugi::xml_node p) {
    RunStyle st;
    for (pugi::xml_node r : p.children("w:r")) {
        if (pugi::xml_node rpr = r.child("w:rPr")) {
            if (pugi::xml_node b = rpr.child("w:b")) {
                const std::string val = b.attribute("w:val").value();
                st.bold = val.empty() || (val != "0" && val != "false");
            }
            if (pugi::xml_node sz = rpr.child("w:sz")) {
                st.size_pt = sz.attribute("w:val").as_double(22.0) / 2.0;
            }
        }
        break;
    }
    return st;
}

std::string paragraph_plain_text(pugi::xml_node p) {
    std::string out;
    for (pugi::xml_node r : p.children("w:r")) {
        for (pugi::xml_node child : r.children()) {
            const std::string_view name = child.name();
            if (name == "w:t") {
                out += child.text().get();
            } else if (name == "w:tab") {
                out += "    ";
            } else if (name == "w:br" || name == "w:cr") {
                out += "\n";
            }
        }
    }
    return out;
}

// draw_line() draws exactly one line — but a string reaching it isn't
// always actually one line: real Excel-authored chart category/series text
// can carry a literal embedded line break (Alt+Enter inside a cell), e.g.
// a two-line axis label serialized as "<c:v>0,5\n0,6</c:v>". PoDoFo's
// embedded-CID-font encoding path throws PdfErrorCode::InvalidFontData on
// a raw control character in a DrawText() string, which would otherwise
// take down the whole save() over one axis label. Collapse any C0 control
// character to a space here — the one choke point nearly everything this
// renderer draws passes through — rather than chasing every place text
// might originate from.
std::string sanitize_single_line(std::string s) {
    for (char& ch : s) {
        if (static_cast<unsigned char>(ch) < 0x20) ch = ' ';
    }
    return s;
}

RelMap build_rel_map(const PartMap& parts) {
    RelMap out;
    const auto it = parts.find("word/_rels/document.xml.rels");
    if (it == parts.end()) return out;

    pugi::xml_document rels_doc;
    if (!rels_doc.load_buffer(it->second.data(), it->second.size())) return out;

    for (pugi::xml_node rel : rels_doc.child("Relationships").children("Relationship")) {
        const std::string id = rel.attribute("Id").value();
        if (!id.empty()) out[id] = rel.attribute("Target").value();
    }
    return out;
}

struct ImageBlock {
    std::string png_bytes;
    double      width_pt  = 0.0;
    double      height_pt = 0.0;
};

// Only called when the caller (draw_paragraph) has already established the
// drawing isn't a chart. Throws NotImplemented if it's neither a picture nor
// a chart (some DrawingML shape TextFabric never produces itself).
ImageBlock resolve_image(pugi::xml_node drawing, const RelMap& rels, const PartMap& parts) {
    pugi::xml_node blip = find_descendant(drawing, "a:blip");
    if (!blip) {
        throw ReportException(
            ReportError::NotImplemented,
            "native PDF backend: <w:drawing> holds neither a recognized "
            "picture nor a chart");
    }

    const std::string rid = blip.attribute("r:embed").value();
    const auto rel_it = rels.find(rid);
    if (rel_it == rels.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("native PDF backend: dangling image relationship '{}'", rid));
    }

    const std::string part_name = "word/" + rel_it->second;
    const auto part_it = parts.find(part_name);
    if (part_it == parts.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("native PDF backend: missing media part '{}'", part_name));
    }

    pugi::xml_node extent = find_descendant(drawing, "wp:extent");
    constexpr double kDefaultEmu = 9525.0 * 96.0;  // 96px square fallback

    ImageBlock img;
    img.png_bytes = part_it->second;
    img.width_pt  = emu_to_pt(extent.attribute("cx").as_double(kDefaultEmu));
    img.height_pt = emu_to_pt(extent.attribute("cy").as_double(kDefaultEmu));
    return img;
}

// ── Charts (bar/line/area/pie/doughnut — see PLAN.md for radar/ofPie/etc.) ─
//
// Reads the same <c:cat>/<c:val> cache shape merger.cpp already parses for
// setChartValue/setChartData, but only ever reads it (this renderer never
// mutates a chart) — kept as its own small parser here rather than sharing
// merger.cpp's anonymous-namespace helpers, since those are tied to editing
// concerns (category matching by name, cache rewriting) this code doesn't
// need.

enum class ChartKind { Bar, Line, Area, Pie, Doughnut };

struct ChartSeriesData {
    std::string          name;
    std::vector<double>  values;  // index-aligned with ChartData::categories
};

struct ChartData {
    std::vector<std::string>     categories;
    std::vector<ChartSeriesData> series;
};

std::string read_series_name(pugi::xml_node ser) {
    pugi::xml_node pt = ser.child("c:tx").child("c:strRef").child("c:strCache").child("c:pt");
    if (pt) return sanitize_single_line(pt.child_value("c:v"));
    return sanitize_single_line(ser.child("c:tx").child_value("c:v"));  // rare literal <c:tx><c:v> form
}

ChartData read_chart_series_data(pugi::xml_node type_node) {
    ChartData data;
    for (pugi::xml_node ser : type_node.children("c:ser")) {
        ChartSeriesData s;
        s.name = read_series_name(ser);

        if (data.categories.empty()) {
            pugi::xml_node cat = ser.child("c:cat");
            pugi::xml_node cache = cat.child("c:strRef").child("c:strCache");
            if (!cache) cache = cat.child("c:numRef").child("c:numCache");
            for (pugi::xml_node pt : cache.children("c:pt")) {
                // Real Excel-authored categories can carry a literal
                // embedded line break (Alt+Enter in a cell) — collapse it,
                // this renderer draws axis labels as a single line.
                data.categories.emplace_back(sanitize_single_line(pt.child_value("c:v")));
            }
        }

        pugi::xml_node val_cache = ser.child("c:val").child("c:numRef").child("c:numCache");
        for (pugi::xml_node pt : val_cache.children("c:pt")) {
            const int idx = pt.attribute("idx").as_int(-1);
            if (idx < 0) continue;
            if (static_cast<std::size_t>(idx) >= s.values.size()) {
                s.values.resize(static_cast<std::size_t>(idx) + 1, 0.0);
            }
            s.values[static_cast<std::size_t>(idx)] = pt.child("c:v").text().as_double();
        }
        data.series.push_back(std::move(s));
    }
    return data;
}

// Fixed palette cycled per series. Real per-series colors live in
// <c:spPr>/<a:solidFill> in the chart XML — reading and translating full
// DrawingML fill definitions is out of scope for this first pass (see
// PLAN.md); every rendered chart uses this palette regardless of what the
// template author picked in Word/Excel.
const PoDoFo::PdfColor& chart_palette_color(std::size_t index) {
    static const PoDoFo::PdfColor kPalette[] = {
        PoDoFo::PdfColor(0.20, 0.40, 0.80),
        PoDoFo::PdfColor(0.85, 0.35, 0.15),
        PoDoFo::PdfColor(0.25, 0.65, 0.35),
        PoDoFo::PdfColor(0.65, 0.25, 0.65),
        PoDoFo::PdfColor(0.85, 0.65, 0.10),
        PoDoFo::PdfColor(0.25, 0.55, 0.65),
    };
    return kPalette[index % (sizeof(kPalette) / sizeof(kPalette[0]))];
}

struct ResolvedChart {
    ChartKind kind      = ChartKind::Bar;
    double    hole_frac = 0.5;   // only meaningful for Doughnut
    ChartData data;
    double    width_pt  = 0.0;   // from the drawing's <wp:extent>
    double    height_pt = 0.0;
};

// Resolves a <c:chart r:id="..."> reference to its part (the relationship
// lives in the same word/_rels/document.xml.rels as image relationships,
// just under a different Id), classifies the plot type, and reads its
// series/category data plus the drawing's natural (unscaled) size. Throws
// NotImplemented for anything outside bar/line/area/pie/doughnut (their 3D
// variants rendered flat, no projection) — see PLAN.md.
ResolvedChart resolve_chart(pugi::xml_node drawing, pugi::xml_node chart_ref,
                            const RelMap& rels, const PartMap& parts) {
    const std::string rid = chart_ref.attribute("r:id").value();
    const auto rel_it = rels.find(rid);
    if (rel_it == rels.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("native PDF backend: dangling chart relationship '{}'", rid));
    }
    const std::string chart_part_name = "word/" + rel_it->second;
    const auto part_it = parts.find(chart_part_name);
    if (part_it == parts.end()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("native PDF backend: missing chart part '{}'", chart_part_name));
    }

    pugi::xml_document chart_doc;
    if (!chart_doc.load_buffer(part_it->second.data(), part_it->second.size())) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            fmt::format("native PDF backend: chart part '{}' failed to parse", chart_part_name));
    }
    pugi::xml_node plot_area =
        chart_doc.child("c:chartSpace").child("c:chart").child("c:plotArea");

    pugi::xml_node type_node;
    ChartKind kind = ChartKind::Bar;
    bool recognized = false;
    bool supported  = false;
    for (pugi::xml_node child : plot_area.children()) {
        const std::string_view name = child.name();
        if (name == "c:barChart" || name == "c:bar3DChart") {
            type_node = child; kind = ChartKind::Bar; recognized = supported = true;
        } else if (name == "c:lineChart" || name == "c:line3DChart") {
            type_node = child; kind = ChartKind::Line; recognized = supported = true;
        } else if (name == "c:areaChart" || name == "c:area3DChart") {
            type_node = child; kind = ChartKind::Area; recognized = supported = true;
        } else if (name == "c:pieChart" || name == "c:pie3DChart") {
            type_node = child; kind = ChartKind::Pie; recognized = supported = true;
        } else if (name == "c:doughnutChart") {
            type_node = child; kind = ChartKind::Doughnut; recognized = supported = true;
        } else if (name == "c:ofPieChart" || name == "c:radarChart"
                || name == "c:scatterChart" || name == "c:bubbleChart"
                || name == "c:stockChart" || name == "c:surfaceChart"
                || name == "c:surface3DChart") {
            recognized = true;
        }
        if (recognized) break;
    }
    if (!supported) {
        throw ReportException(
            ReportError::NotImplemented,
            "native PDF backend renders bar/line/area/pie/doughnut charts "
            "only so far — pie-of-pie/bar-of-pie, radar, scatter, bubble, "
            "stock, and surface charts are not yet implemented (see "
            "PLAN.md); use the Word/LibreOffice converter path for "
            "templates containing one of those");
    }

    ResolvedChart rc;
    rc.kind = kind;
    rc.data = read_chart_series_data(type_node);
    if (rc.data.categories.empty() || rc.data.series.empty()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "native PDF backend: chart has no readable category/series data");
    }
    if (kind == ChartKind::Doughnut) {
        // <c:holeSize val="50"/> — percentage of outer radius the hole
        // occupies; only meaningful for doughnut, absent on pie.
        rc.hole_frac = type_node.child("c:holeSize").attribute("val").as_double(50.0) / 100.0;
    }

    pugi::xml_node extent = find_descendant(drawing, "wp:extent");
    constexpr double kDefaultChartCx = 9525.0 * 640.0;
    constexpr double kDefaultChartCy = 9525.0 * 380.0;
    rc.width_pt  = emu_to_pt(extent.attribute("cx").as_double(kDefaultChartCx));
    rc.height_pt = emu_to_pt(extent.attribute("cy").as_double(kDefaultChartCy));
    return rc;
}

// Finds a table style's own <w:style> node in styles.xml by id — only the
// direct w:type="table" entries; <w:basedOn> inheritance chains (e.g.
// TableGrid and PlainTable4 both declare w:basedOn="TableNormal") aren't
// followed, so a style whose own <w:tblPr> has no border info is treated
// as "this style deliberately specifies none" rather than climbing to its
// parent's. Confirmed against a real template: Word's built-in "Plain
// Table 4" style has no <w:tblBorders> anywhere in its own definition
// (banding comes from cell shading, not lines) and TableGrid defines every
// side as "single" directly — TableNormal's own (usually empty) borders
// never actually come into play for either in practice.
pugi::xml_node find_table_style(const pugi::xml_document& styles_doc, const std::string& style_id) {
    if (style_id.empty()) return {};
    for (pugi::xml_node style : styles_doc.child("w:styles").children("w:style")) {
        if (std::strcmp(style.attribute("w:type").value(), "table") != 0) continue;
        if (style_id == style.attribute("w:styleId").value()) return style;
    }
    return {};
}

// Whether one border side is visible, resolved the way Word actually does
// it, in priority order:
//   1. A per-cell <w:tcPr>/<w:tcBorders> override — real templates
//      routinely borderless a whole table except for e.g. a single
//      underline under the header row, via exactly this mechanism.
//   2. The table's own <w:tblPr>/<w:tblBorders> — its outer edge value
//      (top/left/bottom/right) for a cell actually on that edge of the
//      table, its "inside" value (insideH/insideV) for every interior
//      row/column boundary.
//   3. The referenced table *style*'s own <w:tblBorders> (see
//      find_table_style()) — same outer/inside distinction. A style's
//      per-region conditional formatting (<w:tblStylePr w:type="firstRow">
//      and friends, which real "Plain Table" styles use for bold text and
//      banded shading, not borders) isn't resolved — out of scope, see
//      PLAN.md.
//   4. With no signal from any of the above: `false` (no visible border)
//      if the table named a style that was actually found — even one
//      that, like Plain Table 4, defines no border information anywhere,
//      relying entirely on shading — since that's a deliberate design, not
//      missing data. Only a table with no style reference at all (this
//      project's own examples, which set <w:tblBorders> directly instead)
//      falls back to `true`, preserving the original default for
//      hand-authored templates that specify neither.
bool cell_border_visible(pugi::xml_node tc, pugi::xml_node tbl_borders,
                         pugi::xml_node style_borders, bool style_resolved,
                         const char* cell_side, const char* table_outer_side,
                         const char* table_inside_side, bool is_outer_edge) {
    if (pugi::xml_node tc_borders = tc.child("w:tcPr").child("w:tcBorders")) {
        if (pugi::xml_node b = tc_borders.child(cell_side)) {
            const std::string val = b.attribute("w:val").value();
            return val != "none" && val != "nil";
        }
    }
    const char* side = is_outer_edge ? table_outer_side : table_inside_side;
    if (tbl_borders) {
        if (pugi::xml_node b = tbl_borders.child(side)) {
            const std::string val = b.attribute("w:val").value();
            return val != "none" && val != "nil";
        }
    }
    if (style_borders) {
        if (pugi::xml_node b = style_borders.child(side)) {
            const std::string val = b.attribute("w:val").value();
            return val != "none" && val != "nil";
        }
    }
    return !style_resolved;
}

struct CellBorderSides { bool top, left, bottom, right; };

CellBorderSides resolve_cell_borders(pugi::xml_node tc, pugi::xml_node tbl_borders,
                                     pugi::xml_node style_borders, bool style_resolved,
                                     bool first_row, bool last_row,
                                     bool first_col, bool last_col) {
    CellBorderSides s;
    s.top    = cell_border_visible(tc, tbl_borders, style_borders, style_resolved,
                                   "w:top",    "w:top",    "w:insideH", first_row);
    s.bottom = cell_border_visible(tc, tbl_borders, style_borders, style_resolved,
                                   "w:bottom", "w:bottom", "w:insideH", last_row);
    s.left   = cell_border_visible(tc, tbl_borders, style_borders, style_resolved,
                                   "w:left",   "w:left",   "w:insideV", first_col);
    s.right  = cell_border_visible(tc, tbl_borders, style_borders, style_resolved,
                                   "w:right",  "w:right",  "w:insideV", last_col);
    return s;
}

// Column widths for a table laid out at `width` points wide. `force_fit`
// rescales the template's own <w:tblGrid> widths (originally computed
// relative to the full page) to exactly fill `width` regardless of
// direction — required for a nested table inside a cell, whose available
// width is a fraction of the page and almost never matches what the
// template's author saw. At the page level (force_fit=false) a table
// narrower than the page keeps its own width and stays left-aligned,
// matching the template's intent; it's only ever scaled *down* if it would
// otherwise overflow the page.
std::vector<double> table_column_widths(pugi::xml_node tbl, double width, bool force_fit) {
    std::vector<double> col_widths;
    if (pugi::xml_node grid = tbl.child("w:tblGrid")) {
        for (pugi::xml_node col : grid.children("w:gridCol")) {
            col_widths.push_back(twips_to_pt(col.attribute("w:w").as_llong(0)));
        }
    }
    if (col_widths.empty()) {
        pugi::xml_node first_tr = tbl.child("w:tr");
        std::size_t ncols = 0;
        for (pugi::xml_node tc : first_tr.children("w:tc")) { (void)tc; ++ncols; }
        ncols = std::max<std::size_t>(ncols, 1);
        col_widths.assign(ncols, width / static_cast<double>(ncols));
        return col_widths;
    }
    double sum = 0.0;
    for (double w : col_widths) sum += w;
    if (sum <= 0.0) {
        col_widths.assign(col_widths.size(), width / static_cast<double>(col_widths.size()));
        return col_widths;
    }
    if (force_fit || sum > width) {
        const double scale = width / sum;
        for (double& w : col_widths) w *= scale;
    }
    return col_widths;
}

// Per-column `col_widths` only lines up 1:1 with a row's actual <w:tc>
// elements when no cell in the row merges columns. A cell carrying
// <w:tcPr>/<w:gridSpan val="N"> consumes N consecutive grid columns, not
// one — confirmed as the real cause of a badly-compressed real-world
// nested table: its immediate parent cell spanned both columns of a
// 2-column outer layout table (gridSpan="2", the full page width) but was
// being sized as only the *first* column, compressing everything nested
// inside it by roughly half. Returns one width per entry in `cells`,
// walking the grid column cursor forward by each cell's own span so a
// later cell's index into `col_widths` is never assumed to equal its
// position in `cells`.
std::vector<double> cell_widths_for_row(const std::vector<pugi::xml_node>& cells,
                                        const std::vector<double>& col_widths) {
    std::vector<double> widths;
    widths.reserve(cells.size());
    std::size_t grid_col = 0;
    for (pugi::xml_node tc : cells) {
        std::size_t span = 1;
        if (pugi::xml_node gs = tc.child("w:tcPr").child("w:gridSpan")) {
            span = std::max<unsigned>(1, gs.attribute("w:val").as_uint(1));
        }
        double w = 0.0;
        for (std::size_t k = 0; k < span && grid_col + k < col_widths.size(); ++k) {
            w += col_widths[grid_col + k];
        }
        if (w <= 0.0 && !col_widths.empty()) w = col_widths.back();
        widths.push_back(w);
        grid_col += span;
    }
    return widths;
}

// Explicit env override (mirrors the TEXTFABRIC_SOFFICE pattern) first, then
// a short list of Unicode TrueType fonts commonly present per platform.
// DejaVu Sans is the priority candidate specifically because it covers
// Cyrillic — TextFabric's own example templates rely on Cyrillic to
// demonstrate UTF-8 support, and Standard 14 PDF fonts (WinAnsi-only) can't
// render it at all.
std::optional<std::string> discover_font_path(bool bold) {
    const char* env = std::getenv(bold ? "TEXTFABRIC_NATIVE_PDF_FONT_BOLD"
                                        : "TEXTFABRIC_NATIVE_PDF_FONT");
    if (env != nullptr && *env != '\0' && std::filesystem::exists(env)) {
        return std::string(env);
    }

    static const char* const kRegular[] = {
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    };
    static const char* const kBold[] = {
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
        "C:\\Windows\\Fonts\\arialbd.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    };

    for (const char* candidate : (bold ? kBold : kRegular)) {
        if (std::filesystem::exists(candidate)) return std::string(candidate);
    }
    return std::nullopt;
}

constexpr double kDefaultTableFontSize = 10.0;
constexpr double kCellPadding          = 4.0;

// Small fixed gap dropped before an image or table block so it doesn't sit
// flush against whatever text paragraph came before it (a heading's own
// line box already reserves space for descenders, not a visual margin).
// Plain paragraph-to-paragraph spacing doesn't need this — line_height
// already reads as comfortable there.
constexpr double kBlockGap = 6.0;

// What a table cell's content resolves to for layout purposes. Real-world
// templates mix these more freely than this models (see PLAN.md) — a cell
// is classified as exactly ONE of these, in priority order: a nested table
// wins over everything else in the cell, then the first picture/chart
// drawing, then plain paragraph text. A cell that mixes a drawing with
// meaningful surrounding text only shows the drawing.
struct CellContent {
    enum class Kind { Empty, Text, Image, Chart, NestedTable } kind = Kind::Empty;
    std::vector<std::string> text_lines;
    ImageBlock      image;
    ResolvedChart   chart;
    pugi::xml_node  nested_tbl;
    double          height = 0.0;  // computed once, reused for measure + draw
};

// Walks one in-memory document and lays it out onto a growing PdfMemDocument,
// page by page. Not reentrant / not thread-safe — one Layout per document.
class Layout {
public:
    Layout(PoDoFo::PdfMemDocument& doc, const PageMetrics& pm,
           PoDoFo::PdfFont& regular, PoDoFo::PdfFont& bold,
           const pugi::xml_document& styles_doc)
        : doc_(doc), pm_(pm), regular_(regular), bold_(bold), styles_doc_(styles_doc) {}

    void new_page() {
        if (page_open_) painter_.FinishDrawing();
        PoDoFo::PdfPage& page = doc_.GetPages().CreatePage(
            PoDoFo::Rect(0, 0, pm_.width_pt, pm_.height_pt));
        painter_.SetCanvas(page);
        cursor_y_  = pm_.height_pt - pm_.margin_top_pt;
        page_open_ = true;
    }

    void finish() {
        if (page_open_) {
            painter_.FinishDrawing();
            page_open_ = false;
        }
    }

    void draw_paragraph(pugi::xml_node p, const RelMap& rels, const PartMap& parts) {
        if (pugi::xml_node drawing = find_descendant(p, "w:drawing")) {
            // A paragraph carrying a picture/chart is treated as that block
            // alone — TextFabric never puts running text and a drawing in
            // the same paragraph itself (setImage clears the run range
            // first), so this covers every document this library produces.
            if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
                draw_chart_block(drawing, chart_ref, rels, parts);
            } else {
                draw_image(resolve_image(drawing, rels, parts));
            }
            return;
        }

        const RunStyle style = first_run_style(p);
        const std::string text = paragraph_plain_text(p);
        PoDoFo::PdfFont& font = style.bold ? bold_ : regular_;

        PoDoFo::PdfTextState state;
        state.Font     = &font;
        state.FontSize = style.size_pt;
        const double line_height =
            std::max(font.GetLineSpacing(state), style.size_pt * 1.15);

        if (text.empty()) {
            ensure_space(line_height);
            cursor_y_ -= line_height;
            return;
        }

        for (const std::string& line : state.SplitTextAsLines(text, content_width())) {
            ensure_space(line_height);
            // Decrement before drawing, same convention draw_table() uses:
            // cursor_y_ tracks the top of the next free line box, not a
            // baseline, so the first line under a heading or an image must
            // drop a full line_height before its glyphs are drawn — otherwise
            // ascenders are placed directly on the previous block's bottom
            // edge and visibly overlap it.
            cursor_y_ -= line_height;
            draw_line(font, style.size_pt, line, pm_.margin_left_pt, cursor_y_);
        }
    }

    // Top-level table: the only one that participates in page-break
    // decisions (checked per row, via the shared cursor_y_). A nested
    // table (reached through a cell's own content) never triggers a page
    // break — see layout_table_rows().
    void draw_table(pugi::xml_node tbl, const RelMap& rels, const PartMap& parts) {
        std::vector<pugi::xml_node> row_nodes;
        for (pugi::xml_node tr : tbl.children("w:tr")) row_nodes.push_back(tr);
        if (row_nodes.empty()) return;

        ensure_space(kBlockGap);
        cursor_y_ -= kBlockGap;

        const std::vector<double> col_widths =
            table_column_widths(tbl, content_width(), /*force_fit*/false);
        const TableBorderSources borders = resolve_table_border_sources(tbl);

        PoDoFo::PdfTextState table_state;
        table_state.Font     = &regular_;
        table_state.FontSize = kDefaultTableFontSize;
        const double line_height =
            std::max(regular_.GetLineSpacing(table_state), kDefaultTableFontSize * 1.15);

        for (std::size_t ri = 0; ri < row_nodes.size(); ++ri) {
            pugi::xml_node tr = row_nodes[ri];
            const bool first_row = (ri == 0);
            const bool last_row  = (ri == row_nodes.size() - 1);

            std::vector<pugi::xml_node> cells;
            for (pugi::xml_node tc : tr.children("w:tc")) cells.push_back(tc);
            const std::vector<double> cell_widths = cell_widths_for_row(cells, col_widths);

            std::vector<CellContent> contents(cells.size());
            double row_height = kDefaultTableFontSize * 1.15 + 2 * kCellPadding;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                contents[i] = classify_cell(cells[i], std::max(cell_widths[i] - 2 * kCellPadding, 1.0),
                                            rels, parts, table_state, line_height);
                row_height = std::max(row_height, contents[i].height + 2 * kCellPadding);
            }

            ensure_space(row_height);
            const double row_top = cursor_y_;

            double x = pm_.margin_left_pt;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                const double col_w = cell_widths[i];
                const CellBorderSides sides = resolve_cell_borders(
                    cells[i], borders.tbl_borders, borders.style_borders, borders.style_resolved,
                    first_row, last_row, i == 0, i == cells.size() - 1);
                draw_cell_border(x, row_top - row_height, col_w, row_height, sides);
                draw_cell_content(contents[i], x + kCellPadding, row_top - kCellPadding,
                                  std::max(col_w - 2 * kCellPadding, 1.0), line_height,
                                  rels, parts);
                x += col_w;
            }
            cursor_y_ = row_top - row_height;
        }
    }

private:
    // Resolves the table-level and style-level border sources for one
    // <w:tbl>, once per table rather than once per cell: its own
    // <w:tblPr>/<w:tblBorders>, and — via <w:tblPr>/<w:tblStyle w:val="X">
    // — style "X"'s own <w:tblBorders> plus whether that style was found
    // at all (see cell_border_visible() for how the two combine).
    struct TableBorderSources {
        pugi::xml_node tbl_borders;
        pugi::xml_node style_borders;
        bool           style_resolved = false;
    };

    TableBorderSources resolve_table_border_sources(pugi::xml_node tbl) {
        TableBorderSources src;
        pugi::xml_node tbl_pr = tbl.child("w:tblPr");
        src.tbl_borders = tbl_pr.child("w:tblBorders");
        const std::string style_id = tbl_pr.child("w:tblStyle").attribute("w:val").value();
        if (pugi::xml_node style = find_table_style(styles_doc_, style_id)) {
            src.style_borders  = style.child("w:tblPr").child("w:tblBorders");
            src.style_resolved = true;
        }
        return src;
    }

    // Classifies one <w:tc> and measures the height its content needs at
    // `width` points wide — a nested <w:tbl> recurses (dry run, no
    // drawing) to get its own total height. Shared between the page-level
    // draw_table() and layout_table_rows() (nested tables) so both use
    // identical sizing.
    CellContent classify_cell(pugi::xml_node tc, double width,
                              const RelMap& rels, const PartMap& parts,
                              PoDoFo::PdfTextState& text_state, double line_height) {
        CellContent c;
        if (pugi::xml_node nested = tc.child("w:tbl")) {
            c.kind = CellContent::Kind::NestedTable;
            c.nested_tbl = nested;
            c.height = layout_table_rows(nested, 0, 0, width, rels, parts, /*dry_run*/true);
            return c;
        }

        for (pugi::xml_node p : tc.children("w:p")) {
            pugi::xml_node drawing = find_descendant(p, "w:drawing");
            if (!drawing) continue;
            if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
                c.kind  = CellContent::Kind::Chart;
                c.chart = resolve_chart(drawing, chart_ref, rels, parts);
                c.height = c.chart.height_pt;
                if (c.chart.width_pt > width && c.chart.width_pt > 0.0) {
                    c.height *= width / c.chart.width_pt;
                }
            } else {
                c.kind  = CellContent::Kind::Image;
                c.image = resolve_image(drawing, rels, parts);
                c.height = c.image.height_pt;
                if (c.image.width_pt > width && c.image.width_pt > 0.0) {
                    c.height *= width / c.image.width_pt;
                }
            }
            return c;
        }

        std::string text;
        for (pugi::xml_node p : tc.children("w:p")) {
            if (!text.empty()) text += "\n";
            text += paragraph_plain_text(p);
        }
        c.kind       = CellContent::Kind::Text;
        c.text_lines = text_state.SplitTextAsLines(text, width);
        c.height     = static_cast<double>(std::max<std::size_t>(c.text_lines.size(), 1)) * line_height;
        return c;
    }

    // Draws only the sides resolve_cell_borders() found visible, as
    // independent line segments rather than one rectangle — two adjacent
    // cells that disagree about their shared edge (a header cell's own
    // <w:tcBorders> claiming a bottom line the row below doesn't repeat as
    // its own top, say) both still get an honest rendering: whichever side
    // claims the line draws it, redrawing an edge both sides claim is
    // harmless.
    void draw_cell_border(double x, double y, double w, double h, const CellBorderSides& sides) {
        // A preceding line/area chart leaves the stroke width/color set to
        // whatever it last drew a series with (draw_chart_plot() doesn't
        // reset those two, only the colors used for fills/text) — pin both
        // explicitly rather than let a border silently inherit that.
        painter_.GraphicsState.SetLineWidth(1.0);
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        if (sides.top)    painter_.DrawLine(x, y + h, x + w, y + h);
        if (sides.bottom) painter_.DrawLine(x, y, x + w, y);
        if (sides.left)   painter_.DrawLine(x, y, x, y + h);
        if (sides.right)  painter_.DrawLine(x + w, y, x + w, y + h);
    }

    // Draws one already-classified cell's content, top-aligned within the
    // row, at the given content-box origin (inside the cell padding — the
    // caller already drew the cell's own border rectangle, if any).
    void draw_cell_content(const CellContent& c, double x, double top_y, double width,
                           double line_height, const RelMap& rels, const PartMap& parts) {
        switch (c.kind) {
            case CellContent::Kind::Text: {
                double ty = top_y;
                for (const std::string& line : c.text_lines) {
                    ty -= line_height;
                    draw_line(regular_, kDefaultTableFontSize, line, x, ty);
                }
                break;
            }
            case CellContent::Kind::Image: {
                std::unique_ptr<PoDoFo::PdfImage> image = doc_.CreateImage();
                image->LoadFromBuffer(
                    PoDoFo::bufferview(c.image.png_bytes.data(), c.image.png_bytes.size()));
                const double h = c.height;
                const double w = (c.image.height_pt > 0.0)
                                     ? c.image.width_pt * (h / c.image.height_pt) : width;
                const double scale_x = w / static_cast<double>(image->GetWidth());
                const double scale_y = h / static_cast<double>(image->GetHeight());
                painter_.DrawImage(*image, x, top_y - h, scale_x, scale_y);
                break;
            }
            case CellContent::Kind::Chart: {
                const double h = c.height;
                const double w = (c.chart.height_pt > 0.0)
                                     ? c.chart.width_pt * (h / c.chart.height_pt) : width;
                // No legend — cells are usually too narrow to spare the
                // extra row, and the category/series names are already
                // visible via the rest of the cell's own table row/column
                // headers in every real template this was tested against.
                if (c.chart.kind == ChartKind::Pie || c.chart.kind == ChartKind::Doughnut) {
                    draw_pie_chart_plot(c.chart.kind == ChartKind::Doughnut, c.chart.hole_frac,
                                        c.chart.data, x, top_y - h, w, h);
                } else {
                    draw_chart_plot(c.chart.kind, c.chart.data, x, top_y - h, w, h);
                }
                break;
            }
            case CellContent::Kind::NestedTable:
                layout_table_rows(c.nested_tbl, x, top_y, width, rels, parts, /*dry_run*/false);
                break;
            case CellContent::Kind::Empty:
                break;
        }
    }

    // A table laid out at a fixed (x, width) with no page-break awareness —
    // used only for a table nested inside another table's cell, which is
    // already confined to whatever vertical space the outer row measured
    // for it. `dry_run` skips every painter_ call and just returns the
    // total height, for the outer cell's own sizing pass. A nested table
    // that overflows the bottom margin is a known limitation (see
    // PLAN.md) — it draws into the margin rather than breaking pages.
    double layout_table_rows(pugi::xml_node tbl, double x, double top_y, double width,
                             const RelMap& rels, const PartMap& parts, bool dry_run) {
        std::vector<pugi::xml_node> row_nodes;
        for (pugi::xml_node tr : tbl.children("w:tr")) row_nodes.push_back(tr);
        if (row_nodes.empty()) return 0.0;

        const std::vector<double> col_widths = table_column_widths(tbl, width, /*force_fit*/true);
        const TableBorderSources borders = resolve_table_border_sources(tbl);

        PoDoFo::PdfTextState table_state;
        table_state.Font     = &regular_;
        table_state.FontSize = kDefaultTableFontSize;
        const double line_height =
            std::max(regular_.GetLineSpacing(table_state), kDefaultTableFontSize * 1.15);

        double y = top_y;
        for (std::size_t ri = 0; ri < row_nodes.size(); ++ri) {
            pugi::xml_node tr = row_nodes[ri];
            const bool first_row = (ri == 0);
            const bool last_row  = (ri == row_nodes.size() - 1);

            std::vector<pugi::xml_node> cells;
            for (pugi::xml_node tc : tr.children("w:tc")) cells.push_back(tc);
            const std::vector<double> cell_widths = cell_widths_for_row(cells, col_widths);

            std::vector<CellContent> contents(cells.size());
            double row_height = kDefaultTableFontSize * 1.15 + 2 * kCellPadding;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                contents[i] = classify_cell(cells[i], std::max(cell_widths[i] - 2 * kCellPadding, 1.0),
                                            rels, parts, table_state, line_height);
                row_height = std::max(row_height, contents[i].height + 2 * kCellPadding);
            }

            if (!dry_run) {
                double cx = x;
                for (std::size_t i = 0; i < cells.size(); ++i) {
                    const double col_w = cell_widths[i];
                    const CellBorderSides sides = resolve_cell_borders(
                        cells[i], borders.tbl_borders, borders.style_borders, borders.style_resolved,
                        first_row, last_row, i == 0, i == cells.size() - 1);
                    draw_cell_border(cx, y - row_height, col_w, row_height, sides);
                    draw_cell_content(contents[i], cx + kCellPadding, y - kCellPadding,
                                      std::max(col_w - 2 * kCellPadding, 1.0), line_height,
                                      rels, parts);
                    cx += col_w;
                }
            }
            y -= row_height;
        }
        return top_y - y;
    }

    double content_width() const {
        return pm_.width_pt - pm_.margin_left_pt - pm_.margin_right_pt;
    }

    void ensure_space(double needed) {
        if (cursor_y_ - needed < pm_.margin_bottom_pt) new_page();
    }

    void draw_line(PoDoFo::PdfFont& font, double size, const std::string& text,
                   double x, double y) {
        // DrawText() manages its own BT/ET text object internally in this
        // PoDoFo version (BeginText/EndText are private — meant for a
        // lower-level multi-call text object API this renderer doesn't need).
        painter_.TextState.SetFont(font, size);
        painter_.DrawText(sanitize_single_line(text), x, y);
    }

    void draw_image(const ImageBlock& img) {
        ensure_space(img.height_pt + kBlockGap);
        cursor_y_ -= kBlockGap;
        std::unique_ptr<PoDoFo::PdfImage> image = doc_.CreateImage();
        image->LoadFromBuffer(PoDoFo::bufferview(img.png_bytes.data(), img.png_bytes.size()));

        const double scale_x = img.width_pt  / static_cast<double>(image->GetWidth());
        const double scale_y = img.height_pt / static_cast<double>(image->GetHeight());
        painter_.DrawImage(*image, pm_.margin_left_pt, cursor_y_ - img.height_pt,
                            scale_x, scale_y);
        cursor_y_ -= img.height_pt;
    }

    // Resolves and renders a chart at page scale (bar/line/area/pie/
    // doughnut, 3D variants rendered flat) or throws NotImplemented —
    // see resolve_chart(). This is the page-flow entry point (with a
    // legend row); a chart inside a table cell goes through
    // draw_cell_content() instead, which skips the legend.
    void draw_chart_block(pugi::xml_node drawing, pugi::xml_node chart_ref,
                          const RelMap& rels, const PartMap& parts) {
        const ResolvedChart rc = resolve_chart(drawing, chart_ref, rels, parts);
        const double width_pt  = std::min(rc.width_pt, content_width());
        const double height_pt = rc.height_pt;

        constexpr double kLegendRowHeight = 16.0;
        ensure_space(height_pt + kLegendRowHeight + kBlockGap);
        cursor_y_ -= kBlockGap;

        const double chart_top = cursor_y_;
        if (rc.kind == ChartKind::Pie || rc.kind == ChartKind::Doughnut) {
            draw_pie_chart_plot(rc.kind == ChartKind::Doughnut, rc.hole_frac, rc.data,
                               pm_.margin_left_pt, chart_top - height_pt,
                               width_pt, height_pt);
            cursor_y_ = chart_top - height_pt;
            draw_category_legend(rc.data, pm_.margin_left_pt, cursor_y_ - 4.0);
        } else {
            draw_chart_plot(rc.kind, rc.data, pm_.margin_left_pt, chart_top - height_pt,
                            width_pt, height_pt);
            cursor_y_ = chart_top - height_pt;
            draw_chart_legend(rc.data, pm_.margin_left_pt, cursor_y_ - 4.0);
        }
        cursor_y_ -= kLegendRowHeight;
    }

    // Pie/doughnut layout: a single series read as one wedge per category
    // (multi-series "concentric ring" doughnuts aren't supported — only
    // data.series.front() is drawn), colored per-category like Word/Excel's
    // <c:varyColors val="1"/> convention rather than per-series. Wedges
    // start at 12 o'clock and sweep clockwise, matching Word/Excel/
    // LibreOffice's own pie orientation.
    void draw_pie_chart_plot(bool doughnut, double hole_frac, const ChartData& data,
                             double x, double y, double w, double h) {
        const auto& values = data.series.front().values;
        const std::size_t n = std::min(values.size(), data.categories.size());

        double total = 0.0;
        for (std::size_t i = 0; i < n; ++i) total += std::max(values[i], 0.0);
        if (total <= 0.0) total = 1.0;

        const double cx = x + w / 2.0;
        const double cy = y + h / 2.0;
        const double radius  = std::min(w, h) / 2.0 * 0.80;
        const double hole_r  = doughnut ? radius * std::clamp(hole_frac, 0.0, 0.9) : 0.0;

        constexpr double kHalfPi = 1.5707963267948966;
        constexpr double kTwoPi  = 6.283185307179586;

        PoDoFo::PdfTextState label_state;
        label_state.Font     = &regular_;
        label_state.FontSize = 8.0;

        double cum = 0.0;
        for (std::size_t i = 0; i < n; ++i) {
            const double frac = std::max(values[i], 0.0) / total;
            if (frac <= 0.0) continue;
            const double a_start = kHalfPi - kTwoPi * cum;
            const double a_end   = kHalfPi - kTwoPi * (cum + frac);
            const double a_mid   = (a_start + a_end) / 2.0;
            cum += frac;

            PoDoFo::PdfPainterPath path;
            if (hole_r > 0.0) {
                path.MoveTo(cx + hole_r * std::cos(a_start), cy + hole_r * std::sin(a_start));
                path.AddLineTo(cx + radius * std::cos(a_start), cy + radius * std::sin(a_start));
                path.AddArc(cx, cy, radius, a_start, a_end, /*clockwise*/false);
                path.AddLineTo(cx + hole_r * std::cos(a_end), cy + hole_r * std::sin(a_end));
                path.AddArc(cx, cy, hole_r, a_end, a_start, /*clockwise*/true);
                path.Close();
            } else {
                path.MoveTo(cx, cy);
                path.AddArc(cx, cy, radius, a_start, a_end, /*clockwise*/false);
                path.Close();
            }
            painter_.GraphicsState.SetNonStrokingColor(chart_palette_color(i));
            painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Fill);

            char pct_buf[16];
            std::snprintf(pct_buf, sizeof(pct_buf), "%.0f%%", frac * 100.0);
            const double label_r = hole_r > 0.0 ? (hole_r + radius) / 2.0 : radius * 0.65;
            const double label_w = regular_.GetStringLength(pct_buf, label_state);
            painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(1, 1, 1));
            draw_line(regular_, 8.0, pct_buf,
                     cx + label_r * std::cos(a_mid) - label_w / 2.0,
                     cy + label_r * std::sin(a_mid) - 3.0);
        }

        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
    }

    // Legend keyed by category (pie/doughnut have one series, many
    // categories — the wedges, not the series, are what's color-coded).
    // Same single-row-only simplification as draw_chart_legend().
    void draw_category_legend(const ChartData& data, double x, double y) {
        constexpr double kSwatch   = 8.0;
        constexpr double kGap      = 6.0;
        constexpr double kFontSize = 8.0;

        PoDoFo::PdfTextState state;
        state.Font     = &regular_;
        state.FontSize = kFontSize;

        double       cursor_x = x;
        const double max_x    = x + content_width();
        for (std::size_t i = 0; i < data.categories.size(); ++i) {
            const std::string& name    = data.categories[i];
            const double        text_w = regular_.GetStringLength(name, state);
            const double        item_w = kSwatch + 4.0 + text_w + kGap * 2.0;
            if (cursor_x + item_w > max_x && cursor_x > x) break;

            painter_.GraphicsState.SetNonStrokingColor(chart_palette_color(i));
            painter_.DrawRectangle(cursor_x, y, kSwatch, kSwatch, PoDoFo::PdfPathDrawMode::Fill);
            painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            draw_line(regular_, kFontSize, name, cursor_x + kSwatch + 4.0, y + 1.0);
            cursor_x += item_w;
        }
    }

    // Draws axes, gridlines with value labels, category labels, and the
    // series marks themselves (bars / a stroked line / a filled area) inside
    // the box (x, y)-(x+w, y+h). Best-effort: axis scaling is linear from
    // min(0, dataMin) to dataMax with a fixed 4 gridlines, no attempt at
    // "nice" rounded tick values the way Excel/LibreOffice pick them.
    void draw_chart_plot(ChartKind kind, const ChartData& data,
                        double x, double y, double w, double h) {
        constexpr double kAxisLabelSize = 7.0;
        constexpr double kLeftPad   = 34.0;
        constexpr double kBottomPad = 16.0;
        constexpr double kTopPad    = 4.0;
        constexpr double kRightPad  = 4.0;

        const double plot_x = x + kLeftPad;
        const double plot_y = y + kBottomPad;
        const double plot_w = std::max(w - kLeftPad - kRightPad, 1.0);
        const double plot_h = std::max(h - kBottomPad - kTopPad, 1.0);

        double data_min = 0.0;
        double data_max = 0.0;
        bool   first = true;
        for (const auto& s : data.series) {
            for (double v : s.values) {
                if (first) { data_min = data_max = v; first = false; }
                data_min = std::min(data_min, v);
                data_max = std::max(data_max, v);
            }
        }
        data_min = std::min(data_min, 0.0);
        if (data_max <= data_min) data_max = data_min + 1.0;

        const auto val_to_y = [&](double v) {
            return plot_y + (v - data_min) / (data_max - data_min) * plot_h;
        };

        PoDoFo::PdfTextState axis_state;
        axis_state.Font     = &regular_;
        axis_state.FontSize = kAxisLabelSize;

        // Gridlines + value-axis labels.
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0.75, 0.75, 0.75));
        painter_.GraphicsState.SetLineWidth(0.5);
        constexpr int kGridSteps = 4;
        for (int i = 0; i <= kGridSteps; ++i) {
            const double v  = data_min + (data_max - data_min) * i / kGridSteps;
            const double gy = val_to_y(v);
            painter_.DrawLine(plot_x, gy, plot_x + plot_w, gy);

            char buf[32];
            std::snprintf(buf, sizeof(buf), "%.4g", v);
            const double label_w = regular_.GetStringLength(buf, axis_state);
            draw_line(regular_, kAxisLabelSize, buf,
                     plot_x - label_w - 4.0, gy - kAxisLabelSize * 0.3);
        }

        // Axis lines, drawn solid/black over the gridlines.
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetLineWidth(1.0);
        painter_.DrawLine(plot_x, plot_y, plot_x, plot_y + plot_h);
        painter_.DrawLine(plot_x, val_to_y(0.0), plot_x + plot_w, val_to_y(0.0));

        const std::size_t ncat   = data.categories.size();
        const double      slot_w = plot_w / static_cast<double>(std::max<std::size_t>(ncat, 1));

        for (std::size_t i = 0; i < ncat; ++i) {
            const double cx      = plot_x + (static_cast<double>(i) + 0.5) * slot_w;
            const double label_w = regular_.GetStringLength(data.categories[i], axis_state);
            draw_line(regular_, kAxisLabelSize, data.categories[i],
                     cx - label_w / 2.0, plot_y - 10.0);
        }

        const std::size_t nseries    = data.series.size();
        const double      baseline_y = val_to_y(0.0);

        if (kind == ChartKind::Bar) {
            const double group_w = slot_w * 0.7;
            const double bar_w   = group_w / static_cast<double>(std::max<std::size_t>(nseries, 1));
            for (std::size_t si = 0; si < nseries; ++si) {
                painter_.GraphicsState.SetNonStrokingColor(chart_palette_color(si));
                const auto& values = data.series[si].values;
                for (std::size_t i = 0; i < ncat; ++i) {
                    const double v  = (i < values.size()) ? values[i] : 0.0;
                    const double vy = val_to_y(v);
                    const double bx = plot_x + static_cast<double>(i) * slot_w
                                     + (slot_w - group_w) / 2.0 + static_cast<double>(si) * bar_w;
                    const double bar_bottom = std::min(vy, baseline_y);
                    const double bar_h      = std::abs(vy - baseline_y);
                    painter_.DrawRectangle(bx, bar_bottom, bar_w * 0.9,
                                           std::max(bar_h, 0.5),
                                           PoDoFo::PdfPathDrawMode::Fill);
                }
            }
        } else {
            // Line and area share the same point layout; area additionally
            // closes the path down to the zero line and fills it. Series
            // are drawn in template order, so a later area series is drawn
            // over — not blended with — an earlier one where they overlap;
            // there's no alpha compositing here (see PLAN.md).
            for (std::size_t si = 0; si < nseries; ++si) {
                const auto& values = data.series[si].values;
                if (values.empty()) continue;

                PoDoFo::PdfPainterPath path;
                for (std::size_t i = 0; i < ncat; ++i) {
                    const double v  = (i < values.size()) ? values[i] : 0.0;
                    const double cx = plot_x + (static_cast<double>(i) + 0.5) * slot_w;
                    const double cy = val_to_y(v);
                    if (i == 0) path.MoveTo(cx, cy);
                    else        path.AddLineTo(cx, cy);
                }

                if (kind == ChartKind::Area) {
                    const double last_cx  = plot_x + (static_cast<double>(ncat) - 0.5) * slot_w;
                    const double first_cx = plot_x + 0.5 * slot_w;
                    path.AddLineTo(last_cx, baseline_y);
                    path.AddLineTo(first_cx, baseline_y);
                    path.Close();
                    painter_.GraphicsState.SetNonStrokingColor(chart_palette_color(si));
                    painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Fill);
                } else {
                    painter_.GraphicsState.SetStrokingColor(chart_palette_color(si));
                    painter_.GraphicsState.SetLineWidth(1.5);
                    painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Stroke);
                }
            }
        }

        // Reset to black so whatever text/lines follow don't inherit a
        // series color.
        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
    }

    // Single-row legend: colored swatch + series name, left to right. v1
    // simplification — series that don't fit in one row are simply omitted
    // rather than wrapped, since bar/line/area templates rarely carry more
    // than a handful of series.
    void draw_chart_legend(const ChartData& data, double x, double y) {
        constexpr double kSwatch   = 8.0;
        constexpr double kGap      = 6.0;
        constexpr double kFontSize = 8.0;

        PoDoFo::PdfTextState state;
        state.Font     = &regular_;
        state.FontSize = kFontSize;

        double       cursor_x = x;
        const double max_x    = x + content_width();
        for (std::size_t i = 0; i < data.series.size(); ++i) {
            const std::string& name    = data.series[i].name;
            const double        text_w = regular_.GetStringLength(name, state);
            const double        item_w = kSwatch + 4.0 + text_w + kGap * 2.0;
            if (cursor_x + item_w > max_x && cursor_x > x) break;

            painter_.GraphicsState.SetNonStrokingColor(chart_palette_color(i));
            painter_.DrawRectangle(cursor_x, y, kSwatch, kSwatch, PoDoFo::PdfPathDrawMode::Fill);
            painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            draw_line(regular_, kFontSize, name, cursor_x + kSwatch + 4.0, y + 1.0);
            cursor_x += item_w;
        }
    }

    PoDoFo::PdfMemDocument&   doc_;
    PageMetrics               pm_;
    PoDoFo::PdfFont&          regular_;
    PoDoFo::PdfFont&          bold_;
    const pugi::xml_document& styles_doc_;
    PoDoFo::PdfPainter        painter_;
    double                    cursor_y_  = 0.0;
    bool                      page_open_ = false;
};

} // namespace

#endif  // TEXTFABRIC_HAVE_PODOFO

void render_native_pdf(const pugi::xml_document& document,
                        const std::unordered_map<std::string, std::string>& parts,
                        const std::filesystem::path& output_path) {
#if !defined(TEXTFABRIC_HAVE_PODOFO)
    (void)document;
    (void)parts;
    (void)output_path;
    throw ReportException(
        ReportError::NotImplemented,
        "native PDF backend not compiled in — rebuild with "
        "-DTEXTFABRIC_ENABLE_NATIVE_PDF=ON (requires PoDoFo)");
#else
    using namespace PoDoFo;

    pugi::xml_node body = document.child("w:document").child("w:body");
    if (!body) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "native PDF backend: word/document.xml has no <w:body>");
    }

    const PageMetrics pm = read_page_metrics(body);
    const RelMap rels    = build_rel_map(parts);

    // Table *style* default borders (see find_table_style()) — best-effort;
    // an absent or unparsable word/styles.xml just leaves this empty, which
    // resolve_table_border_sources() treats as "no style resolved" (same
    // as if the table named a style that genuinely doesn't exist).
    pugi::xml_document styles_doc;
    if (const auto it = parts.find("word/styles.xml"); it != parts.end()) {
        styles_doc.load_buffer(it->second.data(), it->second.size());
    }

    try {
        PdfMemDocument doc;

        const std::optional<std::string> regular_path = discover_font_path(false);
        const std::optional<std::string> bold_path     = discover_font_path(true);

        PdfFont* regular = regular_path
            ? &doc.GetFonts().GetOrCreateFont(*regular_path)
            : &doc.GetFonts().GetStandard14Font(PdfStandard14FontType::Helvetica);
        PdfFont* bold = bold_path
            ? &doc.GetFonts().GetOrCreateFont(*bold_path)
            : (regular_path ? regular
                             : &doc.GetFonts().GetStandard14Font(PdfStandard14FontType::HelveticaBold));

        Layout layout(doc, pm, *regular, *bold, styles_doc);
        layout.new_page();

        for (pugi::xml_node child : body.children()) {
            const std::string_view name = child.name();
            if (name == "w:p") {
                layout.draw_paragraph(child, rels, parts);
            } else if (name == "w:tbl") {
                layout.draw_table(child, rels, parts);
            }
            // w:sectPr (already consumed above) and anything else the
            // library doesn't produce itself: silently skipped.
        }
        layout.finish();

        doc.Save(output_path.string());
    } catch (const ReportException&) {
        throw;
    } catch (const PdfError& e) {
        throw ReportException(
            ReportError::SaveFailed,
            fmt::format("native PDF backend: PoDoFo failed: {}", e.what()));
    }
#endif
}

} // namespace textfabric::docx
