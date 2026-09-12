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

// ── Charts (bar/line/area only — see PLAN.md for pie/doughnut/radar/3D) ────
//
// Reads the same <c:cat>/<c:val> cache shape merger.cpp already parses for
// setChartValue/setChartData, but only ever reads it (this renderer never
// mutates a chart) — kept as its own small parser here rather than sharing
// merger.cpp's anonymous-namespace helpers, since those are tied to editing
// concerns (category matching by name, cache rewriting) this code doesn't
// need.

enum class ChartKind { Bar, Line, Area };

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
    if (pt) return pt.child_value("c:v");
    return ser.child("c:tx").child_value("c:v");  // rare literal <c:tx><c:v> form
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
                data.categories.emplace_back(pt.child_value("c:v"));
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

// Walks one in-memory document and lays it out onto a growing PdfMemDocument,
// page by page. Not reentrant / not thread-safe — one Layout per document.
class Layout {
public:
    Layout(PoDoFo::PdfMemDocument& doc, const PageMetrics& pm,
           PoDoFo::PdfFont& regular, PoDoFo::PdfFont& bold)
        : doc_(doc), pm_(pm), regular_(regular), bold_(bold) {}

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

    void draw_table(pugi::xml_node tbl) {
        std::vector<pugi::xml_node> row_nodes;
        for (pugi::xml_node tr : tbl.children("w:tr")) row_nodes.push_back(tr);
        if (row_nodes.empty()) return;

        ensure_space(kBlockGap);
        cursor_y_ -= kBlockGap;

        std::vector<double> col_widths;
        if (pugi::xml_node grid = tbl.child("w:tblGrid")) {
            for (pugi::xml_node col : grid.children("w:gridCol")) {
                col_widths.push_back(twips_to_pt(col.attribute("w:w").as_llong(0)));
            }
        }
        if (col_widths.empty()) {
            std::size_t ncols = 0;
            for (auto tc : row_nodes.front().children("w:tc")) { (void)tc; ++ncols; }
            ncols = std::max<std::size_t>(ncols, 1);
            col_widths.assign(ncols, content_width() / static_cast<double>(ncols));
        }

        PoDoFo::PdfTextState table_state;
        table_state.Font     = &regular_;
        table_state.FontSize = kDefaultTableFontSize;
        const double line_height =
            std::max(regular_.GetLineSpacing(table_state), kDefaultTableFontSize * 1.15);

        for (pugi::xml_node tr : row_nodes) {
            std::vector<pugi::xml_node> cells;
            for (pugi::xml_node tc : tr.children("w:tc")) cells.push_back(tc);

            std::vector<std::vector<std::string>> cell_lines(cells.size());
            std::size_t max_lines = 1;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                std::string text;
                for (pugi::xml_node p : cells[i].children("w:p")) {
                    if (!text.empty()) text += "\n";
                    text += paragraph_plain_text(p);
                }
                const double col_w = (i < col_widths.size() ? col_widths[i] : col_widths.back());
                cell_lines[i] = table_state.SplitTextAsLines(
                    text, std::max(col_w - 2 * kCellPadding, 1.0));
                max_lines = std::max(max_lines, cell_lines[i].size());
            }

            const double row_height = static_cast<double>(max_lines) * line_height + 2 * kCellPadding;
            ensure_space(row_height);
            const double row_top = cursor_y_;

            double x = pm_.margin_left_pt;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                const double col_w = (i < col_widths.size() ? col_widths[i] : col_widths.back());
                painter_.DrawRectangle(x, row_top - row_height, col_w, row_height);

                double ty = row_top - kCellPadding;
                for (const std::string& line : cell_lines[i]) {
                    ty -= line_height;
                    draw_line(regular_, kDefaultTableFontSize, line, x + kCellPadding, ty);
                }
                x += col_w;
            }
            cursor_y_ = row_top - row_height;
        }
    }

private:
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
        painter_.DrawText(text, x, y);
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

    // Resolves the chart part via the drawing's <c:chart r:id="..."> (the
    // relationship lives in the same word/_rels/document.xml.rels as image
    // relationships, just under a different Id), classifies its plot type,
    // and either renders it (bar/line/area, including their 3D variants
    // rendered flat) or throws NotImplemented for everything else.
    void draw_chart_block(pugi::xml_node drawing, pugi::xml_node chart_ref,
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
            } else if (name == "c:pieChart" || name == "c:doughnutChart"
                    || name == "c:ofPieChart" || name == "c:pie3DChart"
                    || name == "c:radarChart" || name == "c:scatterChart"
                    || name == "c:bubbleChart" || name == "c:stockChart"
                    || name == "c:surfaceChart" || name == "c:surface3DChart") {
                recognized = true;
            }
            if (recognized) break;
        }
        if (!supported) {
            throw ReportException(
                ReportError::NotImplemented,
                "native PDF backend renders bar/line/area charts only so far "
                "— pie/doughnut/radar/scatter/bubble/stock/surface charts are "
                "not yet implemented (see PLAN.md); use the Word/LibreOffice "
                "converter path for templates containing one of those");
        }

        const ChartData data = read_chart_series_data(type_node);
        if (data.categories.empty() || data.series.empty()) {
            throw ReportException(
                ReportError::CantCopyDocxTemplate,
                "native PDF backend: chart has no readable category/series data");
        }

        pugi::xml_node extent = find_descendant(drawing, "wp:extent");
        constexpr double kDefaultChartCx = 9525.0 * 640.0;
        constexpr double kDefaultChartCy = 9525.0 * 380.0;
        const double width_pt = std::min(
            emu_to_pt(extent.attribute("cx").as_double(kDefaultChartCx)), content_width());
        const double height_pt = emu_to_pt(extent.attribute("cy").as_double(kDefaultChartCy));

        constexpr double kLegendRowHeight = 16.0;
        ensure_space(height_pt + kLegendRowHeight + kBlockGap);
        cursor_y_ -= kBlockGap;

        const double chart_top = cursor_y_;
        draw_chart_plot(kind, data, pm_.margin_left_pt, chart_top - height_pt,
                        width_pt, height_pt);
        cursor_y_ = chart_top - height_pt;
        draw_chart_legend(data, pm_.margin_left_pt, cursor_y_ - 4.0);
        cursor_y_ -= kLegendRowHeight;
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

    PoDoFo::PdfMemDocument& doc_;
    PageMetrics             pm_;
    PoDoFo::PdfFont&        regular_;
    PoDoFo::PdfFont&        bold_;
    PoDoFo::PdfPainter      painter_;
    double                  cursor_y_  = 0.0;
    bool                    page_open_ = false;
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

        Layout layout(doc, pm, *regular, *bold);
        layout.new_page();

        for (pugi::xml_node child : body.children()) {
            const std::string_view name = child.name();
            if (name == "w:p") {
                layout.draw_paragraph(child, rels, parts);
            } else if (name == "w:tbl") {
                layout.draw_table(child);
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
