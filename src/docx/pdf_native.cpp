#include "docx/pdf_native.hpp"
#include "docweft/error.hpp"

#include <fmt/format.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <functional>
#include <map>
#include <memory>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#if defined(DOCWEFT_HAVE_PODOFO)
#include <podofo/podofo.h>
#endif

namespace docweft::docx {

#if defined(DOCWEFT_HAVE_PODOFO)

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
    double header_dist_pt  = 36.0;    // <w:pgMar w:header>: page top → header top
    double footer_dist_pt  = 36.0;    // <w:pgMar w:footer>: page bottom → footer bottom
};

PageMetrics read_page_metrics(pugi::xml_node sect) {
    PageMetrics pm;
    if (pugi::xml_node sz = sect.child("w:pgSz")) {
        pm.width_pt  = twips_to_pt(sz.attribute("w:w").as_llong(11906));
        pm.height_pt = twips_to_pt(sz.attribute("w:h").as_llong(16838));
    }
    if (pugi::xml_node mar = sect.child("w:pgMar")) {
        pm.margin_left_pt   = twips_to_pt(mar.attribute("w:left").as_llong(1440));
        pm.margin_right_pt  = twips_to_pt(mar.attribute("w:right").as_llong(1440));
        pm.margin_top_pt    = twips_to_pt(mar.attribute("w:top").as_llong(1440));
        pm.margin_bottom_pt = twips_to_pt(mar.attribute("w:bottom").as_llong(1440));
        pm.header_dist_pt   = twips_to_pt(mar.attribute("w:header").as_llong(720));
        pm.footer_dist_pt   = twips_to_pt(mar.attribute("w:footer").as_llong(720));
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

// ── Style model ─────────────────────────────────────────────────────────────

struct Color {
    double r = 0.0, g = 0.0, b = 0.0;
    bool operator==(const Color& o) const { return r == o.r && g == o.g && b == o.b; }
};

// "C00000" → color; "auto" / malformed → nullopt.
std::optional<Color> parse_hex_color(std::string_view hex) {
    if (hex.size() != 6) return std::nullopt;
    unsigned v = 0;
    for (const char ch : hex) {
        v <<= 4;
        if (ch >= '0' && ch <= '9')      v |= static_cast<unsigned>(ch - '0');
        else if (ch >= 'a' && ch <= 'f') v |= static_cast<unsigned>(ch - 'a' + 10);
        else if (ch >= 'A' && ch <= 'F') v |= static_cast<unsigned>(ch - 'A' + 10);
        else return std::nullopt;
    }
    return Color{((v >> 16) & 0xFF) / 255.0, ((v >> 8) & 0xFF) / 255.0, (v & 0xFF) / 255.0};
}

// Document theme (word/theme/theme1.xml): color scheme and the major/minor
// fonts that "theme font" references resolve to. Defaults are Office's.
struct Theme {
    std::string major_font = "Calibri Light";
    std::string minor_font = "Calibri";
    std::map<std::string, Color> colors = {
        {"dk1", {0, 0, 0}}, {"lt1", {1, 1, 1}},
        {"dk2", *parse_hex_color("44546A")}, {"lt2", *parse_hex_color("E7E6E6")},
        {"accent1", *parse_hex_color("4472C4")}, {"accent2", *parse_hex_color("ED7D31")},
        {"accent3", *parse_hex_color("A5A5A5")}, {"accent4", *parse_hex_color("FFC000")},
        {"accent5", *parse_hex_color("5B9BD5")}, {"accent6", *parse_hex_color("70AD47")},
        {"hlink", *parse_hex_color("0563C1")}, {"folHlink", *parse_hex_color("954F72")}};

    void load(const pugi::xml_document& theme_doc) {
        pugi::xml_node elements = theme_doc.child("a:theme").child("a:themeElements");
        for (pugi::xml_node c : elements.child("a:clrScheme").children()) {
            std::string name = c.name();
            if (name.rfind("a:", 0) != 0) continue;
            name = name.substr(2);
            if (auto rgb = parse_hex_color(c.child("a:srgbClr").attribute("val").value())) {
                colors[name] = *rgb;
            } else if (auto sys = parse_hex_color(c.child("a:sysClr").attribute("lastClr").value())) {
                colors[name] = *sys;
            }
        }
        pugi::xml_node fonts = elements.child("a:fontScheme");
        if (const char* f = fonts.child("a:majorFont").child("a:latin").attribute("typeface").value(); *f) major_font = f;
        if (const char* f = fonts.child("a:minorFont").child("a:latin").attribute("typeface").value(); *f) minor_font = f;
    }

    std::optional<Color> scheme(std::string_view name) const {
        // Aliases used in DrawingML and WordprocessingML.
        std::string key(name);
        if (key == "tx1" || key == "text1" || key == "dark1") key = "dk1";
        else if (key == "bg1" || key == "background1" || key == "light1") key = "lt1";
        else if (key == "tx2" || key == "text2" || key == "dark2") key = "dk2";
        else if (key == "bg2" || key == "background2" || key == "light2") key = "lt2";
        else if (key == "hyperlink") key = "hlink";
        else if (key == "followedHyperlink") key = "folHlink";
        const auto it = colors.find(key);
        if (it == colors.end()) return std::nullopt;
        return it->second;
    }

    Color accent(std::size_t i) const {
        return *scheme("accent" + std::to_string(i % 6 + 1));
    }

    // Theme font reference ("minorHAnsi", "majorAscii", ...) → family.
    std::string font(std::string_view ref) const {
        return ref.rfind("major", 0) == 0 ? major_font : minor_font;
    }
};

// `col` with its HSL luminance L replaced by L * mod + off.
Color hsl_luminance(Color col, double mod, double off) {
    const double mx = std::max({col.r, col.g, col.b});
    const double mn = std::min({col.r, col.g, col.b});
    double h = 0.0, sat = 0.0;
    const double l = (mx + mn) / 2.0;
    if (mx != mn) {
        const double d = mx - mn;
        sat = l > 0.5 ? d / (2.0 - mx - mn) : d / (mx + mn);
        if (mx == col.r)      h = (col.g - col.b) / d + (col.g < col.b ? 6.0 : 0.0);
        else if (mx == col.g) h = (col.b - col.r) / d + 2.0;
        else                  h = (col.r - col.g) / d + 4.0;
        h /= 6.0;
    }
    const double nl = std::clamp(l * mod + off, 0.0, 1.0);
    auto hue = [](double p2, double q, double t) {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p2 + (q - p2) * 6 * t;
        if (t < 0.5) return q;
        if (t < 2.0 / 3) return p2 + (q - p2) * (2.0 / 3 - t) * 6;
        return p2;
    };
    if (sat == 0.0) return Color{nl, nl, nl};
    const double q  = nl < 0.5 ? nl * (1 + sat) : nl + sat - nl * sat;
    const double p2 = 2 * nl - q;
    return Color{hue(p2, q, h + 1.0 / 3), hue(p2, q, h), hue(p2, q, h - 1.0 / 3)};
}

// A DrawingML color element's children (lumMod/lumOff/tint/shade, in
// 1/1000 %) applied to `c`.
Color apply_color_modifiers(Color c, pugi::xml_node color_elem) {
    double mod = 1.0, off = 0.0;
    for (pugi::xml_node m : color_elem.children()) {
        const double v = m.attribute("val").as_double(100000.0) / 100000.0;
        const std::string_view name = m.name();
        if (name == "a:lumMod") mod = v;
        else if (name == "a:lumOff") off = v;
        else if (name == "a:shade") c = Color{c.r * v, c.g * v, c.b * v};
        else if (name == "a:tint") c = Color{c.r + (1 - c.r) * (1 - v), c.g + (1 - c.g) * (1 - v), c.b + (1 - c.b) * (1 - v)};
    }
    if (mod != 1.0 || off != 0.0) c = hsl_luminance(c, mod, off);
    return c;
}

// A WordprocessingML color: the hex value, which Word writes as its own
// resolved copy of any theme color; without one, the theme color with
// its tint/shade (hex bytes), applied to HSL luminance as Word does —
// shade: L·s, tint: L·t + 1 − t, then cut down to whole bytes (Word
// truncates; this reproduces its palette exactly: accent1 4472C4 darker
// 25 % = 2F5496, lighter 80 % = D9E2F3). nullopt for "auto" or nothing
// usable.
std::optional<Color> word_color(pugi::xml_node n, const char* hex_attr, const char* theme_attr,
                                const char* tint_attr, const char* shade_attr, const Theme& theme) {
    if (auto rgb = parse_hex_color(n.attribute(hex_attr).value())) return rgb;
    const std::optional<Color> base = theme.scheme(n.attribute(theme_attr).value());
    if (!base) return std::nullopt;
    auto fraction = [&](const char* attr) -> std::optional<double> {
        const char* v = n.attribute(attr).value();
        if (*v == '\0') return std::nullopt;
        return static_cast<double>(std::strtoul(v, nullptr, 16) & 0xFF) / 255.0;
    };
    auto to_bytes = [](Color c) {
        auto byte = [](double v) { return std::floor(std::clamp(v, 0.0, 1.0) * 255.0 + 1e-9) / 255.0; };
        return Color{byte(c.r), byte(c.g), byte(c.b)};
    };
    if (auto t = fraction(tint_attr)) return to_bytes(hsl_luminance(*base, *t, 1.0 - *t));
    if (auto sh = fraction(shade_attr)) return to_bytes(hsl_luminance(*base, *sh, 0.0));
    return base;
}

// Fill of a <w:shd>, theme fills included.
std::optional<Color> shading_fill(pugi::xml_node shd, const Theme& theme) {
    return word_color(shd, "w:fill", "w:themeFill", "w:themeFillTint", "w:themeFillShade", theme);
}

// Color of a DrawingML fill/line container (<a:solidFill> inside spPr/ln).
std::optional<Color> drawingml_color(pugi::xml_node solid_fill, const Theme& theme) {
    if (!solid_fill) return std::nullopt;
    for (pugi::xml_node c : solid_fill.children()) {
        const std::string_view name = c.name();
        std::optional<Color> base;
        if (name == "a:srgbClr")        base = parse_hex_color(c.attribute("val").value());
        else if (name == "a:schemeClr") base = theme.scheme(c.attribute("val").value());
        else if (name == "a:sysClr")    base = parse_hex_color(c.attribute("lastClr").value());
        if (base) return apply_color_modifiers(*base, c);
    }
    return std::nullopt;
}

// Resolved character formatting of one run.
struct TextStyle {
    std::string font;       // family; resolved to the document default when empty
    double size_pt   = 11.0;
    bool   bold      = false;
    bool   italic    = false;
    bool   underline = false;
    bool   strike    = false;
    Color  color;           // black unless set
    bool   hidden    = false;  // <w:vanish>: not printed
    bool   caps       = false;  // <w:caps>: drawn in capitals
    bool   small_caps = false;  // <w:smallCaps>: lower case drawn as smaller capitals
    enum class VertAlign { Baseline, Superscript, Subscript } vert_align = VertAlign::Baseline;
    double position_pt = 0.0;   // <w:position>: raised (> 0) or lowered, in points

    // Superscript/subscript: size and baseline shift relative to size_pt.
    // Measured from LibreOffice's rendering of <w:vertAlign>; Word's own
    // values weren't available to compare against.
    static constexpr double kScriptScale    = 0.58;
    static constexpr double kSuperscriptRise = 0.36;
    static constexpr double kSubscriptDrop   = 0.12;

    // Size the glyphs are drawn and measured at.
    double drawn_size() const {
        return vert_align == VertAlign::Baseline ? size_pt : size_pt * kScriptScale;
    }
    // Baseline shift, up positive.
    double rise() const {
        switch (vert_align) {
            case VertAlign::Superscript: return position_pt + size_pt * kSuperscriptRise;
            case VertAlign::Subscript:   return position_pt - size_pt * kSubscriptDrop;
            case VertAlign::Baseline:    break;
        }
        return position_pt;
    }

    bool operator==(const TextStyle& o) const {
        return font == o.font && size_pt == o.size_pt && bold == o.bold && italic == o.italic &&
               underline == o.underline && strike == o.strike && color == o.color &&
               hidden == o.hidden && caps == o.caps && small_caps == o.small_caps &&
               vert_align == o.vert_align && position_pt == o.position_pt;
    }
};

// One border line: a cell side (<w:top>, <w:insideH>, ... of
// <w:tcBorders>/<w:tblBorders>) or a paragraph side (<w:pBdr>).
struct Border {
    enum class Style { None, Single, Double, Dotted, Dashed, DotDash } style = Style::None;
    double width = 0.5;  // points; of each line, for Double
    Color  color;        // black for "auto"
    double space = 0.0;  // paragraph borders: distance from the text, points

    bool visible() const { return style != Style::None; }
    bool operator==(const Border& o) const {
        return style == o.style && width == o.width && color == o.color && space == o.space;
    }
    // For deciding which of two borders on one edge wins.
    double weight() const { return style == Style::Double ? width * 3.0 : width; }
};

// <w:val> line styles map to the nearest drawable one: the double-line
// family (double, triple, thinThick..., thickThin...) to Double; dotted,
// dashed and dot-dash kinds to theirs; everything else (wave, 3-D, art
// borders) to a single line.
Border parse_border(pugi::xml_node b, const Theme& theme) {
    Border out;
    const std::string_view val = b.attribute("w:val").value();
    if (val.empty() || val == "none" || val == "nil") return out;
    auto starts = [&](std::string_view prefix) { return val.substr(0, prefix.size()) == prefix; };
    if (starts("double") || starts("triple") || starts("thinThick") || starts("thickThin")) {
        out.style = Border::Style::Double;
    } else if (val == "dotted") {
        out.style = Border::Style::Dotted;
    } else if (val == "dashed" || val == "dashSmallGap") {
        out.style = Border::Style::Dashed;
    } else if (val == "dotDash" || val == "dotDotDash") {
        out.style = Border::Style::DotDash;
    } else {
        out.style = Border::Style::Single;
    }
    if (pugi::xml_attribute sz = b.attribute("w:sz")) {
        out.width = std::max(sz.as_double(4.0) / 8.0, 0.25);  // eighths of a point
    }
    if (auto c = word_color(b, "w:color", "w:themeColor", "w:themeTint", "w:themeShade", theme)) out.color = *c;
    out.space = b.attribute("w:space").as_double(0.0);
    return out;
}

// Of two borders on one shared edge, the one drawn (ECMA-376 17.4.66,
// simplified): a visible one over none, then the heavier, then the darker;
// on a tie `a`.
const Border& stronger(const Border& a, const Border& b) {
    if (!b.visible()) return a;
    if (!a.visible()) return b;
    if (a.weight() != b.weight()) return a.weight() > b.weight() ? a : b;
    auto brightness = [](const Color& c) { return c.r + c.b + 2.0 * c.g; };
    return brightness(b.color) < brightness(a.color) ? b : a;
}

// One border side of a cell, resolved the way Word actually does it, in
// priority order:
//   1. A per-cell <w:tcPr>/<w:tcBorders> override — real templates
//      routinely borderless a whole table except for e.g. a single
//      underline under the header row, via exactly this mechanism.
//   2. The table style's conditional formatting for the cell's region
//      (<w:tblStylePr>/<w:tcPr>/<w:tcBorders>, e.g. a line under the
//      header row), resolved by Layout::cell_style() into `conditional`.
//   3. The table's own <w:tblPr>/<w:tblBorders> — its outer edge value
//      (top/left/bottom/right) for a cell actually on that edge of the
//      table, its "inside" value (insideH/insideV) for every interior
//      row/column boundary.
//   4. The referenced table *style*'s own <w:tblBorders> (see
//      find_table_style()) — same outer/inside distinction.
//   5. With no signal from any of the above: no border if the table named
//      a style that was actually found — even one that, like Plain Table
//      4, defines no border information anywhere, relying entirely on
//      shading — since that's a deliberate design, not missing data. Only
//      a table with no style reference at all (this project's own
//      examples, which set <w:tblBorders> directly instead) falls back to
//      a thin black line, preserving the original default for
//      hand-authored templates that specify neither.
// Resolved paragraph formatting.
struct ParaProps {
    enum class Align { Left, Center, Right, Justify };
    enum class LineRule { Auto, Exact, AtLeast };
    struct Tab {
        enum class Kind { Left, Center, Right } kind = Kind::Left;
        double pos = 0.0;     // from the text area's left edge
        std::string leader;   // <w:leader>: character repeated across the gap, "" = none
    };

    Align    align      = Align::Left;
    double   ind_left   = 0.0;
    double   ind_right  = 0.0;
    double   ind_first  = 0.0;   // first line offset: > 0 first-line indent, < 0 hanging
    double   before     = 0.0;
    double   after      = 0.0;
    LineRule line_rule  = LineRule::Auto;
    double   line       = 1.0;   // Auto: multiple of single; otherwise points
    bool     contextual = false;
    bool     keep_next  = false;  // <w:keepNext>: on the same page as the next block
    std::vector<Tab> tabs;
    std::string style_id;

    // <w:pBdr> (between: drawn between paragraphs of one bordered group)
    // and <w:shd>.
    struct Borders {
        Border top, left, bottom, right, between;
        bool operator==(const Borders& o) const {
            return top == o.top && left == o.left && bottom == o.bottom && right == o.right &&
                   between == o.between;
        }
    } borders;
    std::optional<Color> shading;

    bool boxed() const {
        return borders.top.visible() || borders.left.visible() || borders.bottom.visible() ||
               borders.right.visible() || borders.between.visible() || shading.has_value();
    }
    // Consecutive paragraphs with the same borders, shading and indents
    // share one box, as in Word.
    bool same_box(const ParaProps& o) const {
        return boxed() && borders == o.borders && shading == o.shading &&
               ind_left == o.ind_left && ind_right == o.ind_right;
    }
};

// ── Document structure helpers ──────────────────────────────────────────────
//
// Word wraps content in several transparent containers that carry no
// rendering of their own: content controls (<w:sdt>/<w:sdtContent>) and
// custom XML markup (<w:customXml>) — at block level, around table rows,
// around cells and inside paragraphs. Every place that walks structure goes
// through these helpers, so content inside those wrappers is never skipped.

bool is_transparent_wrapper(std::string_view name) {
    return name == "w:sdt" || name == "w:sdtContent" || name == "w:customXml";
}

// <w:p>/<w:tbl> children of a block container (body, cell, header/footer,
// content control), in document order.
void collect_blocks(pugi::xml_node container, std::vector<pugi::xml_node>& out) {
    for (pugi::xml_node child : container.children()) {
        const std::string_view name = child.name();
        if (name == "w:p" || name == "w:tbl") out.push_back(child);
        else if (is_transparent_wrapper(name)) collect_blocks(child, out);
    }
}

std::vector<pugi::xml_node> blocks_of(pugi::xml_node container) {
    std::vector<pugi::xml_node> out;
    collect_blocks(container, out);
    return out;
}

void collect_named(pugi::xml_node container, const char* tag, std::vector<pugi::xml_node>& out) {
    for (pugi::xml_node child : container.children()) {
        const std::string_view name = child.name();
        if (name == tag) out.push_back(child);
        else if (is_transparent_wrapper(name)) collect_named(child, tag, out);
    }
}

std::vector<pugi::xml_node> rows_of(pugi::xml_node tbl) {
    std::vector<pugi::xml_node> out;
    collect_named(tbl, "w:tr", out);
    return out;
}

std::vector<pugi::xml_node> cells_of(pugi::xml_node tr) {
    std::vector<pugi::xml_node> out;
    collect_named(tr, "w:tc", out);
    return out;
}

bool on_off(pugi::xml_node n) {
    if (!n) return false;
    const std::string val = n.attribute("w:val").value();
    return val.empty() || (val != "0" && val != "false" && val != "off");
}

// Subtrees that never contribute running text to the paragraph they sit
// in: property blocks, deleted text (tracked changes), drawings and legacy
// VML (a text box's own text is not paragraph text — strict mode rejects
// text boxes, see find_unsupported()), and the mc:Fallback copy of an
// mc:AlternateContent (the mc:Choice branch is the one rendered).
bool skipped_in_text(std::string_view name) {
    return name == "w:pPr" || name == "w:rPr" || name == "w:del" || name == "w:moveFrom"
        || name == "w:drawing" || name == "w:pict" || name == "w:object"
        || name == "mc:Fallback" || name == "w:sectPr";
}

// ── Number formats ──────────────────────────────────────────────────────────
//
// List numbers (<w:numFmt>) and page numbers (<w:pgNumType w:fmt>) share
// Word's ST_NumberFormat names.

std::string roman(int v) {
    static constexpr std::pair<int, const char*> kTable[] = {
        {1000, "m"}, {900, "cm"}, {500, "d"}, {400, "cd"}, {100, "c"}, {90, "xc"},
        {50, "l"}, {40, "xl"}, {10, "x"}, {9, "ix"}, {5, "v"}, {4, "iv"}, {1, "i"}};
    std::string out;
    for (const auto& [n, s] : kTable) {
        while (v >= n) { out += s; v -= n; }
    }
    return out;
}

// Letter sequences repeat the letter once more per cycle, like Word:
// a..z, aa..zz, aaa..
std::string letters(int v, const std::vector<std::string>& alphabet) {
    if (v <= 0) return "0";
    const int n = static_cast<int>(alphabet.size());
    const int repeat = (v - 1) / n + 1;
    std::string out;
    for (int i = 0; i < repeat; ++i) out += alphabet[static_cast<std::size_t>((v - 1) % n)];
    return out;
}

std::string format_number(int v, std::string_view fmt) {
    static const std::vector<std::string> kLatinLower = [] {
        std::vector<std::string> a;
        for (char ch = 'a'; ch <= 'z'; ++ch) a.emplace_back(1, ch);
        return a;
    }();
    static const std::vector<std::string> kLatinUpper = [] {
        std::vector<std::string> a;
        for (char ch = 'A'; ch <= 'Z'; ++ch) a.emplace_back(1, ch);
        return a;
    }();
    // Word's Russian sequence skips ё, й, ъ, ы, ь.
    static const std::vector<std::string> kRuLower = {
        "а", "б", "в", "г", "д", "е", "ж", "з", "и", "к", "л", "м", "н", "о", "п",
        "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "э", "ю", "я"};
    static const std::vector<std::string> kRuUpper = {
        "А", "Б", "В", "Г", "Д", "Е", "Ж", "З", "И", "К", "Л", "М", "Н", "О", "П",
        "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Э", "Ю", "Я"};

    if (fmt == "lowerLetter")  return letters(v, kLatinLower);
    if (fmt == "upperLetter")  return letters(v, kLatinUpper);
    if (fmt == "russianLower") return letters(v, kRuLower);
    if (fmt == "russianUpper") return letters(v, kRuUpper);
    if (fmt == "lowerRoman")   return roman(v);
    if (fmt == "upperRoman") {
        std::string r = roman(v);
        for (char& ch : r) ch = static_cast<char>(std::toupper(static_cast<unsigned char>(ch)));
        return r;
    }
    if (fmt == "decimalZero" && v >= 0 && v < 10) return "0" + std::to_string(v);
    if (fmt == "numberInDash") return "- " + std::to_string(v) + " -";
    return std::to_string(v);  // decimal, and the fallback for anything else
}

// Page numbers for PAGE / NUMPAGES / SECTIONPAGES / PAGEREF fields; 0 or
// no entry = not known yet, the field's cached result (whatever Word last
// computed) is used instead.
// A page's number as its section displays it (<w:pgNumType w:fmt>).
struct PageNumber {
    int         number = 0;
    std::string format;  // "" = decimal
    bool operator==(const PageNumber&) const = default;
};

struct PageFields {
    int page          = 0;
    int total         = 0;
    int section_total = 0;
    std::string page_format;  // the current page's section's number format
    const std::map<std::string, PageNumber>* bookmark_pages = nullptr;  // PAGEREF targets
};

// Field instruction keyword: " PAGE  \* MERGEFORMAT " → "PAGE".
std::string field_keyword(std::string_view instr) {
    std::size_t i = 0;
    while (i < instr.size() && std::isspace(static_cast<unsigned char>(instr[i]))) ++i;
    std::string word;
    while (i < instr.size() && !std::isspace(static_cast<unsigned char>(instr[i]))) {
        word.push_back(static_cast<char>(std::toupper(static_cast<unsigned char>(instr[i]))));
        ++i;
    }
    return word;
}

// Second word of a field instruction: " PAGEREF _Toc123 \\h " → "_Toc123".
std::string field_argument(std::string_view instr) {
    std::size_t i = 0;
    auto skip_space = [&] { while (i < instr.size() && std::isspace(static_cast<unsigned char>(instr[i]))) ++i; };
    skip_space();
    while (i < instr.size() && !std::isspace(static_cast<unsigned char>(instr[i]))) ++i;
    skip_space();
    std::string arg;
    const bool quoted = i < instr.size() && instr[i] == '"';
    if (quoted) ++i;
    while (i < instr.size() && (quoted ? instr[i] != '"' : !std::isspace(static_cast<unsigned char>(instr[i])))) {
        arg.push_back(instr[i++]);
    }
    return arg;
}

// Text for a PAGE/NUMPAGES/SECTIONPAGES/PAGEREF field, or nullopt to keep
// the cached result (unknown keyword, or the number isn't known yet).
// A field's numeric format switch (\* roman, \* ALPHABETIC, ...) as a
// number format name; nullopt without one. MERGEFORMAT and the case
// switches (Upper, Caps, ...) aren't number formats.
std::optional<std::string> field_number_format(std::string_view instr) {
    for (std::size_t at = instr.find("\\*"); at != std::string_view::npos; at = instr.find("\\*", at + 2)) {
        std::size_t i = at + 2;
        while (i < instr.size() && std::isspace(static_cast<unsigned char>(instr[i]))) ++i;
        std::size_t j = i;
        while (j < instr.size() && !std::isspace(static_cast<unsigned char>(instr[j]))) ++j;
        const std::string_view sw = instr.substr(i, j - i);
        // Word: all lower case → lower-case numerals, otherwise upper case.
        if (sw == "roman")                         return "lowerRoman";
        if (sw == "Roman" || sw == "ROMAN")        return "upperRoman";
        if (sw == "alphabetic")                    return "lowerLetter";
        if (sw == "Alphabetic" || sw == "ALPHABETIC") return "upperLetter";
        if (sw == "Arabic" || sw == "arabic")      return "decimal";
        if (sw == "ArabicDash")                    return "numberInDash";
    }
    return std::nullopt;
}

// Text for a PAGE/NUMPAGES/SECTIONPAGES/PAGEREF field, or nullopt to keep
// the cached result (unknown keyword, or the number isn't known yet).
// PAGE and PAGEREF show the number the way the page's section numbers its
// pages; NUMPAGES and SECTIONPAGES are plain numbers, as in Word; a
// field's own \* switch wins over both.
std::optional<std::string> page_field_text(std::string_view instr, const PageFields& pf) {
    const std::string keyword = field_keyword(instr);
    const std::optional<std::string> own = field_number_format(instr);
    auto shown = [&](int n, std::string_view section_format) {
        return format_number(n, own ? std::string_view(*own) : section_format);
    };
    if (keyword == "PAGE" && pf.page > 0) return shown(pf.page, pf.page_format);
    if (keyword == "NUMPAGES" && pf.total > 0) return shown(pf.total, "");
    if (keyword == "SECTIONPAGES" && pf.section_total > 0) return shown(pf.section_total, "");
    if (keyword == "PAGEREF" && pf.bookmark_pages != nullptr) {
        const auto it = pf.bookmark_pages->find(field_argument(instr));
        if (it != pf.bookmark_pages->end()) return shown(it->second.number, it->second.format);
    }
    return std::nullopt;
}

// ── UTF-8 and letter case ───────────────────────────────────────────────────

// Code point starting at text[i]; advances i past it.
char32_t next_code_point(std::string_view text, std::size_t& i) {
    const auto b0 = static_cast<unsigned char>(text[i]);
    char32_t cp = b0;
    std::size_t len = 1;
    if (b0 >= 0xF0)      { len = 4; cp = b0 & 0x07; }
    else if (b0 >= 0xE0) { len = 3; cp = b0 & 0x0F; }
    else if (b0 >= 0xC0) { len = 2; cp = b0 & 0x1F; }
    for (std::size_t k = 1; k < len && i + k < text.size(); ++k) {
        cp = (cp << 6) | (static_cast<unsigned char>(text[i + k]) & 0x3F);
    }
    i += len;
    return cp;
}

void append_utf8(std::string& out, char32_t cp) {
    if (cp < 0x80) {
        out.push_back(static_cast<char>(cp));
    } else if (cp < 0x800) {
        out.push_back(static_cast<char>(0xC0 | (cp >> 6)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else if (cp < 0x10000) {
        out.push_back(static_cast<char>(0xE0 | (cp >> 12)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else {
        out.push_back(static_cast<char>(0xF0 | (cp >> 18)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    }
}

// Upper case of `cp` for <w:caps>/<w:smallCaps>: Latin (incl. Latin-1,
// Extended-A and Vietnamese), Greek and Cyrillic, one code point to one
// (so ß stays ß); anything else is returned unchanged.
char32_t upper_case(char32_t cp) {
    // Ranges where an upper case letter is directly followed by its lower case one.
    auto pair = [&](char32_t first, char32_t last, bool upper_even) -> std::optional<char32_t> {
        if (cp < first || cp > last) return std::nullopt;
        const bool even = cp % 2 == 0;
        return even == upper_even ? cp : cp - 1;
    };
    if (cp >= 'a' && cp <= 'z') return cp - 0x20;
    if (cp < 0xB5) return cp;
    if (cp >= 0xE0 && cp <= 0xFE && cp != 0xF7) return cp - 0x20;
    if (cp == 0xFF)  return 0x178;
    if (cp == 0x131) return 'I';  // dotless i
    for (auto [first, last, upper_even] : {std::tuple{0x100u, 0x137u, true}, std::tuple{0x139u, 0x148u, false},
                                           std::tuple{0x14Au, 0x177u, true}, std::tuple{0x179u, 0x17Eu, false}}) {
        if (auto u = pair(first, last, upper_even)) return *u;
    }
    // Greek
    if (cp >= 0x3B1 && cp <= 0x3C9) return cp == 0x3C2 ? 0x3A3 : cp - 0x20;  // final sigma → Σ
    if (cp == 0x3AC) return 0x386;
    if (cp >= 0x3AD && cp <= 0x3AF) return cp - 0x25;
    if (cp == 0x3CA || cp == 0x3CB) return cp - 0x20;
    if (cp == 0x3CC) return 0x38C;
    if (cp == 0x3CD || cp == 0x3CE) return cp - 0x3F;
    // Cyrillic
    if (cp >= 0x430 && cp <= 0x44F) return cp - 0x20;
    if (cp >= 0x450 && cp <= 0x45F) return cp - 0x50;  // ё, ђ, є, і, ї, ј, љ, њ, ћ, ў, џ, ...
    if (cp == 0x4CF) return 0x4C0;
    for (auto [first, last, upper_even] : {std::tuple{0x460u, 0x481u, true}, std::tuple{0x48Au, 0x4BFu, true},
                                           std::tuple{0x4C1u, 0x4CEu, false}, std::tuple{0x4D0u, 0x52Fu, true},
                                           std::tuple{0x1E00u, 0x1E95u, true}, std::tuple{0x1EA0u, 0x1EFFu, true}}) {
        if (auto u = pair(first, last, upper_even)) return *u;
    }
    return cp;
}

std::string upper_case(std::string_view text) {
    std::string out;
    out.reserve(text.size());
    for (std::size_t i = 0; i < text.size();) append_utf8(out, upper_case(next_code_point(text, i)));
    return out;
}

// Size of lower case letters drawn as small capitals (<w:smallCaps>),
// relative to the run's size.
constexpr double kSmallCapsScale = 0.8;

// One piece of a paragraph's inline content in document order.
struct Fragment {
    enum class Kind { Text, Tab, LineBreak, PageBreak } kind = Kind::Text;
    std::string text;
    TextStyle   style;
    bool        small_caps_lower = false;  // style.size_pt already scaled by kSmallCapsScale
};

// Walks a paragraph's inline content in document order, producing styled
// fragments. `style_of_run` resolves a <w:r>'s formatting. Complex fields
// (<w:fldChar begin/separate/end> + <w:instrText>) are tracked on a stack
// so a field's instruction is never printed and PAGE/NUMPAGES can replace
// their cached result.
class ParagraphText {
public:
    using StyleOfRun = std::function<TextStyle(pugi::xml_node run)>;

    ParagraphText(const PageFields& pf, StyleOfRun style_of_run, TextStyle base)
        : pf_(pf), style_of_run_(std::move(style_of_run)), style_(std::move(base)) {}

    std::vector<Fragment> run(pugi::xml_node p) {
        walk(p);
        return std::move(out_);
    }

private:
    struct Field {
        std::string instr;
        bool        in_result = false;
        bool        replaced  = false;  // page number already emitted
    };

    bool printing() const {
        for (const Field& f : fields_) {
            if (!f.in_result || f.replaced) return false;
        }
        return true;
    }

    // Hidden runs are still walked (their fldChar/instrText keep the field
    // stack right), they just produce nothing.
    // Capitals are applied here, before measuring, so lines break on the
    // text as drawn.
    void emit(Fragment::Kind kind, std::string_view text = {}) {
        if (style_.hidden) return;
        if (kind != Fragment::Kind::Text || !(style_.caps || style_.small_caps)) {
            push(kind, text, style_);
            return;
        }
        if (style_.caps) {  // takes precedence over smallCaps
            push(kind, upper_case(text), style_);
            return;
        }
        // Small capitals: each stretch of lower case letters becomes
        // capitals at a smaller size.
        TextStyle small = style_;
        small.size_pt *= kSmallCapsScale;
        std::string piece;
        bool piece_small = false;
        for (std::size_t i = 0; i < text.size();) {
            const char32_t cp = next_code_point(text, i);
            const char32_t up = upper_case(cp);
            const bool is_small = up != cp;
            if (is_small != piece_small && !piece.empty()) {
                push(kind, piece, piece_small ? small : style_, piece_small);
                piece.clear();
            }
            piece_small = is_small;
            append_utf8(piece, up);
        }
        if (!piece.empty()) push(kind, piece, piece_small ? small : style_, piece_small);
    }

    void push(Fragment::Kind kind, std::string_view text, const TextStyle& style,
              bool small_caps_lower = false) {
        if (kind == Fragment::Kind::Text && !out_.empty() &&
            out_.back().kind == Fragment::Kind::Text && out_.back().style == style) {
            out_.back().text += text;
            return;
        }
        out_.push_back({kind, std::string(text), style, small_caps_lower});
    }

    void walk(pugi::xml_node node) {
        for (pugi::xml_node child : node.children()) {
            const std::string_view name = child.name();
            if (skipped_in_text(name)) continue;

            if (name == "w:r") {
                TextStyle saved = style_;
                style_ = style_of_run_(child);
                walk(child);
                style_ = std::move(saved);
            } else if (name == "w:t") {
                if (printing()) emit(Fragment::Kind::Text, child.text().get());
            } else if (name == "w:tab") {
                if (printing()) emit(Fragment::Kind::Tab);
            } else if (name == "w:br") {
                if (printing()) {
                    emit(std::strcmp(child.attribute("w:type").value(), "page") == 0
                             ? Fragment::Kind::PageBreak : Fragment::Kind::LineBreak);
                }
            } else if (name == "w:cr") {
                if (printing()) emit(Fragment::Kind::LineBreak);
            } else if (name == "w:noBreakHyphen") {
                if (printing()) emit(Fragment::Kind::Text, "-");
            } else if (name == "w:instrText") {
                if (!fields_.empty()) fields_.back().instr += child.text().get();
            } else if (name == "w:fldChar") {
                on_fld_char(child.attribute("w:fldCharType").value());
            } else if (name == "w:fldSimple") {
                if (auto text = page_field_text(child.attribute("w:instr").value(), pf_)) {
                    if (printing()) {
                        TextStyle saved = style_;
                        if (pugi::xml_node r = child.child("w:r")) style_ = style_of_run_(r);
                        emit(Fragment::Kind::Text, *text);
                        style_ = std::move(saved);
                    }
                } else {
                    walk(child);  // cached result runs
                }
            } else {
                // Hyperlinks, smart tags, inserted text, content controls,
                // mc:AlternateContent/mc:Choice, ...
                walk(child);
            }
        }
    }

    void on_fld_char(std::string_view type) {
        if (type == "begin") {
            fields_.push_back({});
        } else if (type == "separate" && !fields_.empty()) {
            Field& f = fields_.back();
            f.in_result = true;
            if (auto text = page_field_text(f.instr, pf_)) {
                emit_for_field(*text);
                f.replaced = true;
            }
        } else if (type == "end" && !fields_.empty()) {
            Field f = std::move(fields_.back());
            fields_.pop_back();
            // A field without a cached result (no "separate").
            if (!f.in_result) {
                if (auto text = page_field_text(f.instr, pf_)) emit_for_field(*text);
            }
        }
    }

    // Emits text for the innermost field, if every enclosing field shows its result.
    void emit_for_field(const std::string& text) {
        for (std::size_t i = 0; i + 1 < fields_.size(); ++i) {
            if (!fields_[i].in_result || fields_[i].replaced) return;
        }
        emit(Fragment::Kind::Text, text);
    }

    const PageFields&     pf_;
    StyleOfRun            style_of_run_;
    TextStyle             style_;
    std::vector<Field>    fields_;
    std::vector<Fragment> out_;
};

// Every <w:drawing> of a paragraph that is actually rendered (not the
// mc:Fallback copy), in document order.
void collect_drawings(pugi::xml_node node, std::vector<pugi::xml_node>& out) {
    for (pugi::xml_node child : node.children()) {
        const std::string_view name = child.name();
        if (name == "mc:Fallback" || name == "w:pPr") continue;
        if (name == "w:drawing") out.push_back(child);
        else collect_drawings(child, out);
    }
}

std::vector<pugi::xml_node> drawings_of(pugi::xml_node p) {
    std::vector<pugi::xml_node> out;
    collect_drawings(p, out);
    return out;
}

// First run that carries the paragraph's visible text (inside hyperlinks,
// content controls, ... too).
pugi::xml_node first_run(pugi::xml_node node) {
    for (pugi::xml_node child : node.children()) {
        const std::string_view name = child.name();
        if (skipped_in_text(name)) continue;
        if (name == "w:r") return child;
        if (pugi::xml_node r = first_run(child)) return r;
    }
    return {};
}

// ── Styles ──────────────────────────────────────────────────────────────────
//
// Formatting is resolved the way Word layers it: document defaults
// (<w:docDefaults>) → [table style, inside a table cell] → paragraph style
// chain (<w:pStyle>, or the default paragraph style; following
// <w:basedOn>) → [numbering level, for paragraph properties] → character
// style chain (<w:rStyle>) → direct formatting.

// What a table style gives one cell: the style's own pPr/rPr/tcPr, then
// those of each conditional <w:tblStylePr> that applies to the cell, in the
// order they apply (later wins). See Layout::cell_style().
struct CellStyle {
    std::vector<pugi::xml_node> ppr;
    std::vector<pugi::xml_node> rpr;
    std::vector<pugi::xml_node> tc_pr;
};

class StyleSheet {
public:
    StyleSheet(const pugi::xml_document& styles, const Theme& theme) : theme_(theme) {
        pugi::xml_node root = styles.child("w:styles");
        rpr_defaults_ = root.child("w:docDefaults").child("w:rPrDefault").child("w:rPr");
        ppr_defaults_ = root.child("w:docDefaults").child("w:pPrDefault").child("w:pPr");
        for (pugi::xml_node st : root.children("w:style")) {
            const std::string id = st.attribute("w:styleId").value();
            if (id.empty()) continue;
            by_id_[id] = st;
            const std::string_view is_default = st.attribute("w:default").value();
            if ((is_default == "1" || is_default == "true" || is_default == "on") &&
                std::strcmp(st.attribute("w:type").value(), "paragraph") == 0) {
                default_paragraph_ = id;
            }
        }
    }

    // Formatting of run `r` inside paragraph `p`.
    TextStyle run_style(pugi::xml_node p, pugi::xml_node r, const CellStyle* cell = nullptr) const {
        TextStyle st = paragraph_base(p, cell);
        pugi::xml_node rpr = r.child("w:rPr");
        for (pugi::xml_node s : chain(rpr.child("w:rStyle").attribute("w:val").value())) {
            apply_rpr(s.child("w:rPr"), st);
        }
        apply_rpr(rpr, st);
        return st;
    }

    // Formatting of the paragraph mark (sizes an empty paragraph's line).
    TextStyle mark_style(pugi::xml_node p, const CellStyle* cell = nullptr) const {
        TextStyle st = paragraph_base(p, cell);
        apply_rpr(p.child("w:pPr").child("w:rPr"), st);
        return st;
    }

    // `numbering_level` is the list level's <w:lvl>, if `p` is a list item:
    // its indents sit between the style's and the paragraph's own.
    ParaProps para_props(pugi::xml_node p, pugi::xml_node numbering_level = {},
                         const CellStyle* cell = nullptr) const {
        ParaProps pp;
        pp.style_id = paragraph_style_id(p);
        apply_ppr(ppr_defaults_, pp);
        if (cell != nullptr) {
            for (pugi::xml_node ppr : cell->ppr) apply_ppr(ppr, pp);
        }
        for (pugi::xml_node s : chain(pp.style_id)) apply_ppr(s.child("w:pPr"), pp);
        if (numbering_level) apply_ppr(numbering_level.child("w:pPr"), pp);
        apply_ppr(p.child("w:pPr"), pp);
        return pp;
    }

    // <w:numPr> that applies to `p`: its own, else its style chain's.
    pugi::xml_node num_pr(pugi::xml_node p) const {
        if (pugi::xml_node own = p.child("w:pPr").child("w:numPr")) return own;
        for (pugi::xml_node s : chain(paragraph_style_id(p))) {
            if (pugi::xml_node n = s.child("w:pPr").child("w:numPr")) return n;
        }
        return {};
    }

    // Formatting on top of the defaults from an rPr (e.g. a list level's).
    void apply_rpr(pugi::xml_node rpr, TextStyle& st) const {
        if (!rpr) return;
        if (pugi::xml_node f = rpr.child("w:rFonts")) {
            // hAnsi covers Cyrillic and other non-ASCII Latin-script text.
            for (const char* attr : {"w:hAnsi", "w:ascii"}) {
                if (const char* v = f.attribute(attr).value(); *v) { st.font = v; break; }
            }
            for (const char* attr : {"w:hAnsiTheme", "w:asciiTheme"}) {
                if (const char* v = f.attribute(attr).value(); *v) { st.font = theme_.font(v); break; }
            }
        }
        if (pugi::xml_node n = rpr.child("w:b")) st.bold = on_off(n);
        if (pugi::xml_node n = rpr.child("w:i")) st.italic = on_off(n);
        if (pugi::xml_node n = rpr.child("w:u")) {
            const std::string_view v = n.attribute("w:val").value();
            st.underline = !(v == "none" || v == "0");
        }
        if (pugi::xml_node n = rpr.child("w:strike")) st.strike = on_off(n);
        if (pugi::xml_node n = rpr.child("w:dstrike")) st.strike = st.strike || on_off(n);
        // <w:webHidden> is deliberately not read: it hides text only in Web
        // Layout view, Word prints it (TOC page numbers carry it).
        if (pugi::xml_node n = rpr.child("w:vanish")) st.hidden = on_off(n);
        if (pugi::xml_node n = rpr.child("w:caps")) st.caps = on_off(n);
        if (pugi::xml_node n = rpr.child("w:smallCaps")) st.small_caps = on_off(n);
        if (pugi::xml_node n = rpr.child("w:vertAlign")) {
            const std::string_view v = n.attribute("w:val").value();
            st.vert_align = v == "superscript" ? TextStyle::VertAlign::Superscript
                          : v == "subscript"   ? TextStyle::VertAlign::Subscript
                                               : TextStyle::VertAlign::Baseline;
        }
        if (pugi::xml_node n = rpr.child("w:position")) {
            st.position_pt = n.attribute("w:val").as_double(0.0) / 2.0;  // half-points
        }
        if (pugi::xml_node sz = rpr.child("w:sz")) {
            const double half_points = sz.attribute("w:val").as_double(0.0);
            if (half_points > 0.0) st.size_pt = half_points / 2.0;
        }
        if (pugi::xml_node c = rpr.child("w:color")) {
            const std::string_view v = c.attribute("w:val").value();
            if (auto col = word_color(c, "w:val", "w:themeColor", "w:themeTint", "w:themeShade", theme_)) {
                st.color = *col;
            } else if (v == "auto") {
                st.color = Color{};
            }
        }
    }

private:
    TextStyle paragraph_base(pugi::xml_node p, const CellStyle* cell) const {
        TextStyle st;
        st.font = theme_.minor_font;
        apply_rpr(rpr_defaults_, st);
        if (cell != nullptr) {
            for (pugi::xml_node rpr : cell->rpr) apply_rpr(rpr, st);
        }
        for (pugi::xml_node s : chain(paragraph_style_id(p))) apply_rpr(s.child("w:rPr"), st);
        return st;
    }

    void apply_ppr(pugi::xml_node ppr, ParaProps& pp) const {
        if (!ppr) return;
        if (pugi::xml_node jc = ppr.child("w:jc")) {
            const std::string_view v = jc.attribute("w:val").value();
            if (v == "center")                                      pp.align = ParaProps::Align::Center;
            else if (v == "right" || v == "end")                    pp.align = ParaProps::Align::Right;
            else if (v == "both" || v == "distribute")              pp.align = ParaProps::Align::Justify;
            else                                                    pp.align = ParaProps::Align::Left;
        }
        if (pugi::xml_node ind = ppr.child("w:ind")) {
            auto twips = [&](const char* a, const char* b) -> std::optional<double> {
                for (const char* name : {a, b}) {
                    if (name == nullptr) continue;
                    if (pugi::xml_attribute at = ind.attribute(name)) return twips_to_pt(at.as_llong(0));
                }
                return std::nullopt;
            };
            if (auto v = twips("w:left", "w:start"))  pp.ind_left  = *v;
            if (auto v = twips("w:right", "w:end"))   pp.ind_right = *v;
            if (auto v = twips("w:firstLine", nullptr)) pp.ind_first = *v;
            if (auto v = twips("w:hanging", nullptr))   pp.ind_first = -*v;
        }
        if (pugi::xml_node sp = ppr.child("w:spacing")) {
            if (pugi::xml_attribute a = sp.attribute("w:before")) pp.before = twips_to_pt(a.as_llong(0));
            if (pugi::xml_attribute a = sp.attribute("w:after"))  pp.after  = twips_to_pt(a.as_llong(0));
            if (pugi::xml_attribute a = sp.attribute("w:line")) {
                const std::string_view rule = sp.attribute("w:lineRule").value();
                if (rule == "exact") {
                    pp.line_rule = ParaProps::LineRule::Exact;
                    pp.line      = twips_to_pt(a.as_llong(240));
                } else if (rule == "atLeast") {
                    pp.line_rule = ParaProps::LineRule::AtLeast;
                    pp.line      = twips_to_pt(a.as_llong(240));
                } else {
                    pp.line_rule = ParaProps::LineRule::Auto;
                    pp.line      = a.as_double(240.0) / 240.0;
                }
            }
        }
        if (pugi::xml_node cs = ppr.child("w:contextualSpacing")) pp.contextual = on_off(cs);
        if (pugi::xml_node kn = ppr.child("w:keepNext")) pp.keep_next = on_off(kn);
        if (pugi::xml_node bdr = ppr.child("w:pBdr")) {
            ParaProps::Borders& b = pp.borders;
            for (auto [name, alt, side] : {std::tuple{"w:top", "", &b.top}, std::tuple{"w:left", "w:start", &b.left},
                                           std::tuple{"w:bottom", "", &b.bottom}, std::tuple{"w:right", "w:end", &b.right},
                                           std::tuple{"w:between", "", &b.between}}) {
                pugi::xml_node n = bdr.child(name);
                if (!n && *alt != '\0') n = bdr.child(alt);
                if (n) *side = parse_border(n, theme_);
            }
        }
        if (pugi::xml_node shd = ppr.child("w:shd")) {
            if (auto fill = shading_fill(shd, theme_)) pp.shading = fill;
            else if (std::strcmp(shd.attribute("w:fill").value(), "auto") == 0 ||
                     std::strcmp(shd.attribute("w:val").value(), "nil") == 0) pp.shading.reset();
        }
        if (pugi::xml_node tabs = ppr.child("w:tabs")) {
            for (pugi::xml_node t : tabs.children("w:tab")) {
                const std::string_view v = t.attribute("w:val").value();
                const double pos = twips_to_pt(t.attribute("w:pos").as_llong(0));
                // A stop at the same position replaces the style's one.
                pp.tabs.erase(std::remove_if(pp.tabs.begin(), pp.tabs.end(),
                                             [&](const ParaProps::Tab& x) { return std::abs(x.pos - pos) < 0.5; }),
                              pp.tabs.end());
                if (v == "clear" || v == "bar") continue;  // a bar tab draws a line, it isn't a stop
                ParaProps::Tab tab;
                tab.pos  = pos;
                tab.kind = (v == "right" || v == "end" || v == "decimal") ? ParaProps::Tab::Kind::Right
                         : v == "center"                                  ? ParaProps::Tab::Kind::Center
                                                                          : ParaProps::Tab::Kind::Left;
                const std::string_view leader = t.attribute("w:leader").value();
                tab.leader = leader == "dot"        ? "."
                           : leader == "hyphen"     ? "-"
                           : leader == "underscore" || leader == "heavy" ? "_"
                           : leader == "middleDot"  ? "\u00B7"
                                                    : "";
                pp.tabs.push_back(tab);
            }
            std::sort(pp.tabs.begin(), pp.tabs.end(),
                      [](const ParaProps::Tab& a, const ParaProps::Tab& b) { return a.pos < b.pos; });
        }
    }

    std::string paragraph_style_id(pugi::xml_node p) const {
        const std::string id = p.child("w:pPr").child("w:pStyle").attribute("w:val").value();
        return id.empty() ? default_paragraph_ : id;
    }

    // Base-most first, so later (more specific) styles override.
    std::vector<pugi::xml_node> chain(const std::string& id) const {
        std::vector<pugi::xml_node> out;
        std::string cur = id;
        while (!cur.empty() && out.size() < 32) {  // bound: guards against cycles
            const auto it = by_id_.find(cur);
            if (it == by_id_.end()) break;
            out.push_back(it->second);
            cur = it->second.child("w:basedOn").attribute("w:val").value();
        }
        std::reverse(out.begin(), out.end());
        return out;
    }

    const Theme&                          theme_;
    pugi::xml_node                        rpr_defaults_;
    pugi::xml_node                        ppr_defaults_;
    std::map<std::string, pugi::xml_node> by_id_;
    std::string                           default_paragraph_;
};

// ── Numbering (list markers) ────────────────────────────────────────────────
//
// Produces the marker text ("1.", "a)", "•", "2.1.") and indents for list
// paragraphs from word/numbering.xml, keeping one counter set per <w:num>.
class Numbering {
public:
    struct Marker {
        std::string    text;
        pugi::xml_node level;  // <w:lvl>: its pPr indents and rPr marker formatting
    };

    explicit Numbering(const pugi::xml_document& numbering) {
        pugi::xml_node root = numbering.child("w:numbering");
        for (pugi::xml_node an : root.children("w:abstractNum")) {
            abstract_[an.attribute("w:abstractNumId").as_int(-1)] = an;
        }
        for (pugi::xml_node num : root.children("w:num")) {
            nums_[num.attribute("w:numId").as_int(-1)] = num;
        }
    }

    static bool format_supported(std::string_view fmt) {
        return fmt.empty() || fmt == "decimal" || fmt == "decimalZero" || fmt == "bullet"
            || fmt == "none" || fmt == "lowerLetter" || fmt == "upperLetter"
            || fmt == "lowerRoman" || fmt == "upperRoman"
            || fmt == "russianLower" || fmt == "russianUpper";
    }

    // numFmt of the level `num_pr` points at, if it is one format_supported()
    // rejects; nullopt otherwise (including "not a list at all").
    std::optional<std::string> unsupported_format(pugi::xml_node num_pr) const {
        const pugi::xml_node lvl = level(num_pr);
        if (!lvl) return std::nullopt;
        const std::string fmt = lvl.child("w:numFmt").attribute("w:val").value();
        if (format_supported(fmt)) return std::nullopt;
        return fmt;
    }

    // List counters, to measure ahead without advancing them.
    struct Snapshot { std::map<int, std::array<int, 9>> value; std::map<int, std::array<bool, 9>> started; };
    Snapshot snapshot() const {
        Snapshot sn;
        for (const auto& [id, c] : counters_) { sn.value[id] = c.value; sn.started[id] = c.started; }
        return sn;
    }
    void restore(const Snapshot& sn) {
        counters_.clear();
        for (const auto& [id, v] : sn.value) { counters_[id].value = v; counters_[id].started = sn.started.at(id); }
    }

    // The <w:lvl> `num_pr` points at, if any.
    pugi::xml_node level_of(pugi::xml_node num_pr) const { return level(num_pr); }

    // Advances the list counters and returns the marker for this paragraph,
    // or nullopt if `num_pr` doesn't make it a list item.
    std::optional<Marker> next(pugi::xml_node num_pr) {
        const int num_id = num_pr.child("w:numId").attribute("w:val").as_int(0);
        const int ilvl   = std::clamp(num_pr.child("w:ilvl").attribute("w:val").as_int(0), 0, 8);
        const pugi::xml_node lvl = level(num_pr);
        if (!lvl) return std::nullopt;

        Counters& c = counters_[num_id];
        const auto li = static_cast<std::size_t>(ilvl);
        c.value[li]   = c.started[li] ? c.value[li] + 1 : start_of(num_id, ilvl);
        c.started[li] = true;
        for (std::size_t d = li + 1; d < c.started.size(); ++d) c.started[d] = false;

        Marker m;
        const std::string fmt = lvl.child("w:numFmt").attribute("w:val").value();
        const std::string lvl_text = lvl.child("w:lvlText").attribute("w:val").value();
        if (fmt == "bullet") {
            m.text = bullet_text(lvl_text);
        } else if (fmt != "none") {
            for (std::size_t i = 0; i < lvl_text.size(); ++i) {
                if (lvl_text[i] == '%' && i + 1 < lvl_text.size() &&
                    lvl_text[i + 1] >= '1' && lvl_text[i + 1] <= '9') {
                    const int k = lvl_text[i + 1] - '1';
                    const auto ki = static_cast<std::size_t>(k);
                    const int v = c.started[ki] ? c.value[ki] : start_of(num_id, k);
                    m.text += format_number(v, k == ilvl ? fmt : level_format(num_id, k));
                    ++i;
                } else {
                    m.text.push_back(lvl_text[i]);
                }
            }
        }
        m.level = lvl;
        return m;
    }

private:
    struct Counters {
        std::array<int, 9>  value{};
        std::array<bool, 9> started{};
    };

    pugi::xml_node level(pugi::xml_node num_pr) const {
        const int num_id = num_pr.child("w:numId").attribute("w:val").as_int(0);
        const int ilvl   = std::clamp(num_pr.child("w:ilvl").attribute("w:val").as_int(0), 0, 8);
        return level(num_id, ilvl);
    }

    pugi::xml_node level(int num_id, int ilvl) const {
        const auto num = nums_.find(num_id);
        if (num_id == 0 || num == nums_.end()) return {};
        // A <w:lvlOverride> may carry a whole replacement <w:lvl>.
        for (pugi::xml_node ov : num->second.children("w:lvlOverride")) {
            if (ov.attribute("w:ilvl").as_int(-1) == ilvl) {
                if (pugi::xml_node l = ov.child("w:lvl")) return l;
            }
        }
        const auto an = abstract_.find(num->second.child("w:abstractNumId").attribute("w:val").as_int(-1));
        if (an == abstract_.end()) return {};
        for (pugi::xml_node l : an->second.children("w:lvl")) {
            if (l.attribute("w:ilvl").as_int(-1) == ilvl) return l;
        }
        return {};
    }

    std::string level_format(int num_id, int ilvl) const {
        return level(num_id, ilvl).child("w:numFmt").attribute("w:val").value();
    }

    int start_of(int num_id, int ilvl) const {
        const auto num = nums_.find(num_id);
        if (num != nums_.end()) {
            for (pugi::xml_node ov : num->second.children("w:lvlOverride")) {
                if (ov.attribute("w:ilvl").as_int(-1) == ilvl) {
                    if (pugi::xml_node so = ov.child("w:startOverride")) {
                        return so.attribute("w:val").as_int(1);
                    }
                }
            }
        }
        return level(num_id, ilvl).child("w:start").attribute("w:val").as_int(1);
    }

    // Symbol/Wingdings bullets are private-use code points (U+F0xx) that only
    // mean something in those fonts — draw them as a plain bullet.
    static std::string bullet_text(const std::string& lvl_text) {
        std::string out;
        for (std::size_t i = 0; i < lvl_text.size(); ++i) {
            const auto b0 = static_cast<unsigned char>(lvl_text[i]);
            if (b0 == 0xEF && i + 2 < lvl_text.size() &&
                (static_cast<unsigned char>(lvl_text[i + 1]) & 0xFC) == 0x80) {
                out += "•";
                i += 2;
            } else {
                out.push_back(lvl_text[i]);
            }
        }
        return out.empty() ? std::string("•") : out;
    }

    std::map<int, pugi::xml_node> abstract_;
    std::map<int, pugi::xml_node> nums_;
    std::map<int, Counters>       counters_;
};

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

// Relationship Id → Target for one part's .rels file (the main document's
// by default; headers and footers have their own).
RelMap build_rel_map(const PartMap& parts,
                     const std::string& rels_part = "word/_rels/document.xml.rels") {
    RelMap out;
    const auto it = parts.find(rels_part);
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
// a chart (some DrawingML shape DocWeft never produces itself).
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

// DrawingML boolean (<c:marker val="1"/>): present without val means true.
bool on_off_attr(pugi::xml_node n) {
    if (!n) return false;
    const std::string_view v = n.attribute("val").value();
    return v.empty() || v == "1" || v == "true";
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
enum class Grouping { Clustered, Stacked, PercentStacked };

// Data labels (<c:dLbls>): which parts a label shows, where, how the
// value is formatted and the text's look.
struct DataLabels {
    bool        value = false, percent = false, category = false, series = false;
    std::string separator = ", ";
    std::string position;     // <c:dLblPos>: outEnd, inEnd, ctr, inBase, bestFit, t, b, l, r; "" = default
    std::string format;       // number format code; "" = the values' own (source-linked) one
    std::string custom_text;  // a point's own <c:tx> text, shown instead
    double      size = 7.0;
    bool        bold = false;
    Color       color = {0.25, 0.25, 0.25};  // Office's default label text (tx1, 75 %)

    bool shown() const { return value || percent || category || series || !custom_text.empty(); }
};

struct ChartSeriesData {
    std::string          name;
    std::vector<double>  values;        // index-aligned with ChartData::categories
    Color                color;         // from the template, else a theme accent
    std::vector<Color>   point_colors;  // per category: pie/doughnut wedges
    ChartKind            kind   = ChartKind::Bar;
    bool                 marker = false;  // line: draw point markers
    std::string          number_format;   // the values' <c:formatCode>
    DataLabels           labels;          // for every point, unless overridden
    std::map<std::size_t, DataLabels> point_labels;  // <c:dLbl> by point index
};

// One chart-type element of the plot area (<c:barChart>, <c:lineChart>,
// ...). A combo chart has several; each may use the secondary value axis.
struct ChartGroup {
    ChartKind                kind      = ChartKind::Bar;
    Grouping                 grouping  = Grouping::Clustered;
    bool                     secondary = false;
    double                   gap       = 1.5;  // bars: <c:gapWidth> / 100
    std::vector<std::size_t> series;           // indices into ChartData::series
};

struct ChartData {
    std::vector<std::string>     categories;
    std::vector<ChartSeriesData> series;
    std::vector<ChartGroup>      groups;
};

std::string read_series_name(pugi::xml_node ser) {
    pugi::xml_node pt = ser.child("c:tx").child("c:strRef").child("c:strCache").child("c:pt");
    if (pt) return sanitize_single_line(pt.child_value("c:v"));
    return sanitize_single_line(ser.child("c:tx").child_value("c:v"));  // rare literal <c:tx><c:v> form
}

// Appends `type_node`'s series (and a group for them) to `data`.
// <c:dLbls> (or one point's <c:dLbl>) on top of `base`: elements it
// doesn't have keep the base's settings; <c:delete> hides the labels.
DataLabels read_data_labels(pugi::xml_node node, DataLabels base, const Theme& theme) {
    if (!node) return base;
    if (on_off_attr(node.child("c:delete"))) {
        DataLabels none;
        return none;
    }
    for (auto [tag, flag] : {std::pair{"c:showVal", &base.value}, std::pair{"c:showPercent", &base.percent},
                             std::pair{"c:showCatName", &base.category}, std::pair{"c:showSerName", &base.series}}) {
        if (pugi::xml_node n = node.child(tag)) *flag = on_off_attr(n);
    }
    if (pugi::xml_node sep = node.child("c:separator")) base.separator = sep.child_value();
    if (pugi::xml_node pos = node.child("c:dLblPos")) base.position = pos.attribute("val").value();
    if (pugi::xml_node nf = node.child("c:numFmt"); nf && std::strcmp(nf.attribute("sourceLinked").value(), "1") != 0) {
        base.format = nf.attribute("formatCode").value();
    }
    pugi::xml_node def = node.child("c:txPr").child("a:p").child("a:pPr").child("a:defRPr");
    if (pugi::xml_attribute sz = def.attribute("sz")) base.size = sz.as_double(700.0) / 100.0;
    if (pugi::xml_attribute b = def.attribute("b")) base.bold = std::strcmp(b.value(), "1") == 0;
    if (auto c = drawingml_color(def.child("a:solidFill"), theme)) base.color = *c;
    // A point's own text (<c:tx><c:rich>), e.g. a hand-edited label.
    std::string text;
    for (pugi::xml_node para : node.child("c:tx").child("c:rich").children("a:p")) {
        for (pugi::xml_node r : para.children("a:r")) text += r.child_value("a:t");
    }
    if (!text.empty()) base.custom_text = sanitize_single_line(text);
    return base;
}

void read_chart_group(pugi::xml_node type_node, ChartKind kind, const Theme& theme, ChartData& data) {
    const bool line_chart = kind == ChartKind::Line;
    ChartGroup group;
    group.kind = kind;
    const std::string_view grouping = type_node.child("c:grouping").attribute("val").value();
    group.grouping = grouping == "stacked"        ? Grouping::Stacked
                   : grouping == "percentStacked" ? Grouping::PercentStacked
                                                  : Grouping::Clustered;
    if (pugi::xml_node gap = type_node.child("c:gapWidth")) group.gap = gap.attribute("val").as_double(150.0) / 100.0;
    const bool group_markers = on_off_attr(type_node.child("c:marker"));

    for (pugi::xml_node ser : type_node.children("c:ser")) {
        ChartSeriesData s;
        s.name = read_series_name(ser);
        s.kind = kind;
        if (line_chart) {
            pugi::xml_node marker = ser.child("c:marker");
            const std::string_view symbol = marker.child("c:symbol").attribute("val").value();
            s.marker = marker ? symbol != "none" : group_markers;
        }

        // A line series' color is its stroke; everything else is filled.
        pugi::xml_node sp_pr = ser.child("c:spPr");
        std::optional<Color> color = line_chart
            ? drawingml_color(sp_pr.child("a:ln").child("a:solidFill"), theme)
            : drawingml_color(sp_pr.child("a:solidFill"), theme);
        if (!color) color = drawingml_color(sp_pr.child("a:solidFill"), theme);
        if (!color) color = drawingml_color(sp_pr.child("a:ln").child("a:solidFill"), theme);
        // Default colors follow the series' <c:idx>, as Office assigns them.
        s.color = color.value_or(theme.accent(ser.child("c:idx").attribute("val").as_uint(
            static_cast<unsigned>(data.series.size()))));

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
        s.number_format = val_cache.child_value("c:formatCode");
        // Labels: the group's <c:dLbls>, the series' own over it, then
        // per point.
        const DataLabels group_labels = read_data_labels(type_node.child("c:dLbls"), DataLabels{}, theme);
        pugi::xml_node ser_labels = ser.child("c:dLbls");
        s.labels = read_data_labels(ser_labels, group_labels, theme);
        for (pugi::xml_node dl : ser_labels.children("c:dLbl")) {
            const auto idx = dl.child("c:idx").attribute("val").as_uint(~0u);
            if (idx != ~0u) s.point_labels[idx] = read_data_labels(dl, s.labels, theme);
        }
        for (pugi::xml_node pt : val_cache.children("c:pt")) {
            const int idx = pt.attribute("idx").as_int(-1);
            if (idx < 0) continue;
            if (static_cast<std::size_t>(idx) >= s.values.size()) {
                s.values.resize(static_cast<std::size_t>(idx) + 1, 0.0);
            }
            s.values[static_cast<std::size_t>(idx)] = pt.child("c:v").text().as_double();
        }
        // Per-point colors (<c:dPt>), else the accent sequence — what a
        // varyColors chart (pie/doughnut) shows for each category.
        const std::size_t points = std::max(s.values.size(), data.categories.size());
        for (std::size_t k = 0; k < points; ++k) s.point_colors.push_back(theme.accent(k));
        for (pugi::xml_node dpt : ser.children("c:dPt")) {
            const auto idx = dpt.child("c:idx").attribute("val").as_uint(~0u);
            if (idx >= s.point_colors.size()) continue;
            if (auto c = drawingml_color(dpt.child("c:spPr").child("a:solidFill"), theme)) {
                s.point_colors[idx] = *c;
            }
        }
        group.series.push_back(data.series.size());
        data.series.push_back(std::move(s));
    }
    data.groups.push_back(std::move(group));
}

PoDoFo::PdfColor pdf_color(const Color& c) { return PoDoFo::PdfColor(c.r, c.g, c.b); }

// Wedge `i` of a pie/doughnut (the first series' per-point colors).
Color point_color(const ChartData& data, std::size_t i) {
    const auto& colors = data.series.front().point_colors;
    return i < colors.size() ? colors[i] : Color{0.5, 0.5, 0.5};
}

// Arc of radius `r` around (cx, cy) from angle `a0` to `a1` (radians,
// counter-clockwise from +x; a1 < a0 sweeps clockwise), as cubic Béziers of
// at most 90° each. The path's current point must already be the arc's
// start. Built by hand instead of PdfPainterPath::AddArc: its handling of
// direction and of negative / decreasing angles differs between PoDoFo
// versions, and pie slices came out scrambled with 1.1.2.
void append_arc(PoDoFo::PdfPainterPath& path, double cx, double cy, double r,
                double a0, double a1) {
    constexpr double kQuarter = 1.5707963267948966;
    const int segments = std::max(1, static_cast<int>(std::ceil(std::abs(a1 - a0) / kQuarter - 1e-9)));
    const double delta = (a1 - a0) / segments;
    // Control point distance for a circular arc of angle delta.
    const double k = 4.0 / 3.0 * std::tan(delta / 4.0) * r;
    for (int i = 0; i < segments; ++i) {
        const double t0 = a0 + delta * i;
        const double t1 = t0 + delta;
        const double x0 = cx + r * std::cos(t0), y0 = cy + r * std::sin(t0);
        const double x3 = cx + r * std::cos(t1), y3 = cy + r * std::sin(t1);
        path.AddCubicBezierTo(x0 - k * std::sin(t0), y0 + k * std::cos(t0),
                              x3 + k * std::sin(t1), y3 - k * std::cos(t1),
                              x3, y3);
    }
}

// Value axis scale the way spreadsheet charts pick it: always includes 0,
// about five intervals of a "nice" step (1, 2, 2.5 or 5 × 10^n), ends
// rounded out to whole steps — 0..25000 by 5000 for data up to 21080.
struct AxisScale {
    double min      = 0.0;
    double max      = 1.0;
    double step     = 1.0;
    int    decimals = 0;  // digits after the point that the step needs
};

AxisScale nice_axis_scale(double lo, double hi) {
    lo = std::min(lo, 0.0);
    hi = std::max(hi, 0.0);
    if (hi - lo <= 0.0) hi = lo + 1.0;

    const double raw = (hi - lo) / 5.0;
    const double mag = std::pow(10.0, std::floor(std::log10(raw)));
    double step = 10.0 * mag;
    for (const double m : {1.0, 2.0, 2.5, 5.0}) {
        if (raw <= m * mag * (1.0 + 1e-9)) { step = m * mag; break; }
    }

    AxisScale sc;
    sc.step = step;
    sc.min  = std::floor(lo / step + 1e-9) * step;
    sc.max  = std::ceil(hi / step - 1e-9) * step;
    if (sc.max <= sc.min) sc.max = sc.min + step;
    while (sc.decimals < 6) {
        const double scaled = step * std::pow(10.0, sc.decimals);
        if (std::abs(scaled - std::round(scaled)) < 1e-6 * std::max(1.0, scaled)) break;
        ++sc.decimals;
    }
    return sc;
}

std::string format_axis_value(double v, int decimals) {
    if (std::abs(v) < 1e-12) v = 0.0;  // no "-0"
    return fmt::format("{:.{}f}", v, decimals);
}

// A chart value in an Excel number format code, the kinds data labels
// use: General; fixed decimals ("0", "0.00"); thousands grouping
// ("#,##0"); percent ("0%", "0.0%"); literal text around the number
// ("0 \"₽\"", "\$#,##0"). Only the first section of "pos;neg;zero" codes
// is used (negatives keep their minus sign); [colors]/[$-locale] are
// skipped. Separators are "." and "," like the axis labels.
std::string format_chart_number(double v, std::string_view code) {
    code = code.substr(0, code.find(';'));
    std::string prefix, suffix, pattern;
    for (std::size_t i = 0; i < code.size(); ++i) {
        const char ch = code[i];
        std::string literal;
        if (code.substr(i, 7) == "General") {
            pattern += 'G';
            i += 6;
            continue;
        }
        if (ch == '"') {
            const std::size_t end = code.find('"', i + 1);
            literal = std::string(code.substr(i + 1, end == std::string_view::npos ? std::string_view::npos : end - i - 1));
            i = end == std::string_view::npos ? code.size() : end;
        } else if (ch == '\\' && i + 1 < code.size()) {
            literal = std::string(1, code[++i]);
        } else if (ch == '[') {
            const std::size_t end = code.find(']', i);
            i = end == std::string_view::npos ? code.size() : end;
            continue;
        } else if (ch == '_' || ch == '*') {
            ++i;  // padding / fill character
            continue;
        } else if (std::strchr("0#?.,%", ch) != nullptr) {
            pattern += ch;
            continue;
        } else {
            literal = std::string(1, ch);
        }
        (pattern.empty() ? prefix : suffix) += literal;
    }
    if (pattern.empty() || pattern.find('G') != std::string::npos) {
        if (std::abs(v) < 1e-12) return "0";
        std::string g = fmt::format("{:.10g}", v);
        if (g.find('e') != std::string::npos) g = fmt::format("{:.2f}", v);
        return prefix + g + suffix;
    }
    const bool percent = pattern.find('%') != std::string::npos;
    const std::size_t dot = pattern.find('.');
    int decimals = 0;
    if (dot != std::string::npos) {
        for (std::size_t i = dot + 1; i < pattern.size(); ++i) {
            if (pattern[i] == '0' || pattern[i] == '#' || pattern[i] == '?') ++decimals;
        }
    }
    const std::string_view int_part = std::string_view(pattern).substr(0, dot);
    const bool grouping = int_part.find(',') != std::string_view::npos &&
                          int_part.find_first_of("0#?", int_part.find(',')) != std::string_view::npos;
    double x = percent ? v * 100.0 : v;
    // Excel rounds halves away from zero (4.25 → "4.3"), fmt to even.
    const double scale = std::pow(10.0, decimals);
    x = std::round(x * scale * (1.0 + 1e-12)) / scale;
    if (std::abs(x) < 0.5 / scale) x = 0.0;  // no "-0"
    std::string num = fmt::format("{:.{}f}", std::abs(x), decimals);
    if (grouping) {
        const std::size_t int_end = num.find('.') == std::string::npos ? num.size() : num.find('.');
        for (std::size_t k = int_end; k > 3; k -= 3) num.insert(k - 3, ",");
    }
    return (x < 0.0 ? "-" : "") + prefix + num + (percent ? "%" : "") + suffix;
}

// Labels for point `i` of a series: its own <c:dLbl>, else the series'.
const DataLabels& labels_at(const ChartSeriesData& s, std::size_t i) {
    const auto it = s.point_labels.find(i);
    return it != s.point_labels.end() ? it->second : s.labels;
}

// A data label's text: series name, category, value, percentage — in
// Office's order, joined by the separator. `fraction`: the point's share
// of the whole (pie/doughnut), for the percentage.
std::string data_label_text(const DataLabels& dl, const ChartSeriesData& s, const std::string& category,
                            double value, std::optional<double> fraction = std::nullopt) {
    if (!dl.custom_text.empty()) return dl.custom_text;
    std::vector<std::string> parts;
    if (dl.series) parts.push_back(s.name);
    if (dl.category) parts.push_back(category);
    if (dl.value) parts.push_back(format_chart_number(value, dl.format.empty() ? std::string_view(s.number_format) : std::string_view(dl.format)));
    if (dl.percent && fraction) {
        parts.push_back(format_chart_number(*fraction, dl.format.find('%') != std::string::npos ? std::string_view(dl.format) : "0%"));
    }
    std::string out;
    for (const std::string& part : parts) {
        if (!out.empty()) out += dl.separator;
        out += part;
    }
    return sanitize_single_line(out);
}

struct ChartAxis {
    std::string           title;
    std::optional<double> min, max;        // <c:scaling> overrides
    bool                  reversed = false;  // <c:orientation val="maxMin">
    bool                  deleted  = false;
};

struct ResolvedChart {
    ChartKind kind      = ChartKind::Bar;  // Pie/Doughnut, or Bar for any axis chart
    double    hole_frac = 0.5;   // only meaningful for Doughnut
    ChartData data;
    std::string title;           // empty = no title shown
    bool      horizontal = false;  // bar chart with <c:barDir val="bar">
    bool      has_secondary = false;
    ChartAxis cat_axis;
    ChartAxis val_axis;
    ChartAxis val2_axis;         // secondary value axis
    double    width_pt  = 0.0;   // from the drawing's <wp:extent>
    double    height_pt = 0.0;
};

// DrawingML boolean (<c:delete val="1"/>): present without val means true.
bool on_off_val(pugi::xml_node n) {
    if (!n) return false;
    const std::string val = n.attribute("val").value();
    return val.empty() || val == "1" || val == "true";
}

// Plain text of a <c:title>'s rich text (paragraphs joined by a space).
std::string rich_title_text(pugi::xml_node title) {
    std::string out;
    for (pugi::xml_node para : title.child("c:tx").child("c:rich").children("a:p")) {
        std::string line;
        for (pugi::xml_node r : para.children()) {
            if (std::strcmp(r.name(), "a:r") == 0 || std::strcmp(r.name(), "a:fld") == 0) {
                line += r.child_value("a:t");
            }
        }
        if (line.empty()) continue;
        if (!out.empty()) out += ' ';
        out += line;
    }
    return sanitize_single_line(out);
}

// Chart title as Word shows it: the title's own text; for an auto title
// (<c:title> without text) the series name of a single-series chart.
std::string chart_title_text(pugi::xml_node chart, const ChartData& data) {
    pugi::xml_node title = chart.child("c:title");
    if (!title) return {};
    if (std::string text = rich_title_text(title); !text.empty()) return text;
    if (title.child("c:tx").child("c:strRef")) {
        return sanitize_single_line(title.child("c:tx").child("c:strRef").child("c:strCache")
                                        .child("c:pt").child_value("c:v"));
    }
    return data.series.size() == 1 ? data.series.front().name : std::string{};
}

// Resolves a <c:chart r:id="..."> reference to its part (the relationship
// lives in the same word/_rels/document.xml.rels as image relationships,
// just under a different Id), classifies the plot type, and reads its
// series/category data plus the drawing's natural (unscaled) size. Throws
// NotImplemented for anything outside bar/line/area/pie/doughnut (their 3D
// variants rendered flat, no projection) — see PLAN.md.
ResolvedChart resolve_chart(pugi::xml_node drawing, pugi::xml_node chart_ref,
                            const RelMap& rels, const PartMap& parts, const Theme& theme) {
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

    ResolvedChart rc;
    std::vector<std::pair<pugi::xml_node, ChartKind>> types;
    for (pugi::xml_node child : plot_area.children()) {
        const std::string_view name = child.name();
        if (name == "c:barChart" || name == "c:bar3DChart")          types.emplace_back(child, ChartKind::Bar);
        else if (name == "c:lineChart" || name == "c:line3DChart")   types.emplace_back(child, ChartKind::Line);
        else if (name == "c:areaChart" || name == "c:area3DChart")   types.emplace_back(child, ChartKind::Area);
        else if (name == "c:pieChart" || name == "c:pie3DChart")     types.emplace_back(child, ChartKind::Pie);
        else if (name == "c:doughnutChart")                          types.emplace_back(child, ChartKind::Doughnut);
        else if (name == "c:ofPieChart" || name == "c:radarChart" || name == "c:scatterChart" ||
                 name == "c:bubbleChart" || name == "c:stockChart" || name == "c:surfaceChart" ||
                 name == "c:surface3DChart") {
            throw ReportException(
                ReportError::NotImplemented,
                "native PDF backend renders bar/line/area/pie/doughnut charts "
                "only so far — pie-of-pie/bar-of-pie, radar, scatter, bubble, "
                "stock, and surface charts are not yet implemented (see "
                "PLAN.md); use the Word/LibreOffice converter path for "
                "templates containing one of those");
        }
    }
    if (types.empty()) {
        throw ReportException(ReportError::CantCopyDocxTemplate,
                              "native PDF backend: chart has no plot");
    }
    const bool pie = types.front().second == ChartKind::Pie || types.front().second == ChartKind::Doughnut;
    if (pie && types.size() > 1) {
        throw ReportException(ReportError::NotImplemented,
                              "native PDF backend: pie charts combined with other chart types");
    }

    // Axes by id; the first group's value axis is the primary one.
    std::map<std::string, pugi::xml_node> axes;
    for (pugi::xml_node ax : plot_area.children()) {
        const std::string_view name = ax.name();
        if (name == "c:catAx" || name == "c:dateAx" || name == "c:valAx" || name == "c:serAx") {
            axes[ax.child("c:axId").attribute("val").value()] = ax;
        }
    }
    auto axis_of = [&](pugi::xml_node type_node, const char* axis_tag) -> pugi::xml_node {
        for (pugi::xml_node id : type_node.children("c:axId")) {
            const auto it = axes.find(id.attribute("val").value());
            if (it != axes.end() && std::strcmp(it->second.name(), axis_tag) == 0) return it->second;
        }
        return {};
    };
    auto read_axis = [&](pugi::xml_node ax) {
        ChartAxis a;
        if (!ax) return a;
        a.title    = rich_title_text(ax.child("c:title"));
        a.deleted  = on_off_val(ax.child("c:delete"));
        a.reversed = std::strcmp(ax.child("c:scaling").child("c:orientation").attribute("val").value(), "maxMin") == 0;
        if (pugi::xml_node mn = ax.child("c:scaling").child("c:min")) a.min = mn.attribute("val").as_double();
        if (pugi::xml_node mx = ax.child("c:scaling").child("c:max")) a.max = mx.attribute("val").as_double();
        return a;
    };

    pugi::xml_node primary_val;
    for (const auto& [type_node, kind] : types) {
        read_chart_group(type_node, kind, theme, rc.data);
        if (pie) continue;
        pugi::xml_node val_ax = axis_of(type_node, "c:valAx");
        if (!primary_val) {
            // Groups normally name their axes by id; fall back to the
            // plot area's first axes for charts that don't.
            if (!val_ax) val_ax = plot_area.child("c:valAx");
            primary_val = val_ax;
            pugi::xml_node cat_ax = axis_of(type_node, "c:catAx");
            if (!cat_ax) cat_ax = axis_of(type_node, "c:dateAx");
            if (!cat_ax) cat_ax = plot_area.child("c:catAx");
            if (!cat_ax) cat_ax = plot_area.child("c:dateAx");
            rc.cat_axis = read_axis(cat_ax);
            rc.val_axis = read_axis(val_ax);
        } else if (val_ax && val_ax != primary_val) {
            rc.data.groups.back().secondary = true;
            rc.has_secondary = true;
            rc.val2_axis = read_axis(val_ax);
        }
        if (kind == ChartKind::Bar &&
            std::strcmp(type_node.child("c:barDir").attribute("val").value(), "bar") == 0) {
            rc.horizontal = true;
        }
    }
    if (rc.horizontal && types.size() > 1) {
        throw ReportException(ReportError::NotImplemented,
                              "native PDF backend: horizontal bar charts combined with other chart types");
    }

    rc.kind = pie ? types.front().second : ChartKind::Bar;
    if (rc.data.categories.empty() || rc.data.series.empty()) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "native PDF backend: chart has no readable category/series data");
    }
    rc.title = chart_title_text(chart_doc.child("c:chartSpace").child("c:chart"), rc.data);
    if (rc.kind == ChartKind::Doughnut) {
        // <c:holeSize val="50"/> — percentage of outer radius the hole
        // occupies; only meaningful for doughnut, absent on pie.
        rc.hole_frac = types.front().first.child("c:holeSize").attribute("val").as_double(50.0) / 100.0;
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

// A border side by name; left/right also by their bidi-neutral names.
pugi::xml_node border_side(pugi::xml_node borders, const char* side) {
    if (pugi::xml_node b = borders.child(side)) return b;
    if (std::strcmp(side, "w:left") == 0)  return borders.child("w:start");
    if (std::strcmp(side, "w:right") == 0) return borders.child("w:end");
    return {};
}

Border cell_border(const Theme& theme, pugi::xml_node tc, pugi::xml_node tbl_borders,
                   pugi::xml_node style_borders, bool style_resolved,
                   const std::optional<Border>& conditional,
                   const char* cell_side, const char* table_outer_side,
                   const char* table_inside_side, bool is_outer_edge) {
    if (pugi::xml_node b = border_side(tc.child("w:tcPr").child("w:tcBorders"), cell_side)) return parse_border(b, theme);
    if (conditional) return *conditional;
    const char* side = is_outer_edge ? table_outer_side : table_inside_side;
    if (pugi::xml_node b = border_side(tbl_borders, side)) return parse_border(b, theme);
    if (pugi::xml_node b = border_side(style_borders, side)) return parse_border(b, theme);
    Border fallback;
    if (!style_resolved) fallback.style = Border::Style::Single;
    return fallback;
}

struct CellBorderSides { Border top, left, bottom, right; };

// Borders a table style's conditional formatting gives a cell, per side;
// nullopt = says nothing.
struct StyleBorders {
    std::optional<Border> top, left, bottom, right;
};

CellBorderSides resolve_cell_borders(const Theme& theme, pugi::xml_node tc, pugi::xml_node tbl_borders,
                                     pugi::xml_node style_borders, bool style_resolved,
                                     const StyleBorders& conditional,
                                     bool first_row, bool last_row,
                                     bool first_col, bool last_col) {
    CellBorderSides s;
    s.top    = cell_border(theme, tc, tbl_borders, style_borders, style_resolved, conditional.top,
                           "w:top",    "w:top",    "w:insideH", first_row);
    s.bottom = cell_border(theme, tc, tbl_borders, style_borders, style_resolved, conditional.bottom,
                           "w:bottom", "w:bottom", "w:insideH", last_row);
    s.left   = cell_border(theme, tc, tbl_borders, style_borders, style_resolved, conditional.left,
                           "w:left",   "w:left",   "w:insideV", first_col);
    s.right  = cell_border(theme, tc, tbl_borders, style_borders, style_resolved, conditional.right,
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
        const std::vector<pugi::xml_node> rows = rows_of(tbl);
        std::size_t ncols = rows.empty() ? 0 : cells_of(rows.front()).size();
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

// ── Strict mode ─────────────────────────────────────────────────────────────
//
// Before drawing anything, the whole document (and the headers/footers it
// uses) is checked for content this renderer would silently get wrong —
// dropped text, merged cells drawn as blank ones, a stacked chart drawn as
// a clustered one. Any hit throws NotImplemented listing all of them, so
// save()'s converter chain moves on (or reports a clear error) instead of
// returning a plausible-looking but incomplete PDF. Purely cosmetic gaps
// (fonts, alignment, colors, floating image position) are not rejected.
// DOCWEFT_NATIVE_PDF_STRICT=0 turns the check off.

constexpr const char* kPictureUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
constexpr const char* kChartUri   = "http://schemas.openxmlformats.org/drawingml/2006/chart";

bool strict_mode_enabled() {
    const char* v = std::getenv("DOCWEFT_NATIVE_PDF_STRICT");
    if (v == nullptr) return true;
    const std::string_view val = v;
    return !(val == "0" || val == "false" || val == "off" || val == "no");
}

// Image container by magic bytes; PNG and JPEG are what PoDoFo embeds here.
std::string image_format(const std::string& bytes) {
    auto starts = [&](std::string_view magic) {
        return bytes.size() >= magic.size() && bytes.compare(0, magic.size(), magic) == 0;
    };
    if (starts("\x89PNG")) return "png";
    if (starts("\xFF\xD8\xFF")) return "jpeg";
    if (starts("GIF8")) return "GIF";
    if (starts("BM")) return "BMP";
    if (starts("II*") || starts("MM\x00*")) return "TIFF";
    if (bytes.size() >= 44 && bytes.compare(40, 4, " EMF") == 0) return "EMF";
    if (starts("\xD7\xCD\xC6\x9A") || starts(std::string_view("\x01\x00\x09\x00", 4))) return "WMF";
    if (bytes.find("<svg") != std::string::npos) return "SVG";
    return "unknown";
}

class UnsupportedScan {
public:
    UnsupportedScan(const PartMap& parts, const StyleSheet& styles, const Numbering& numbering)
        : parts_(parts), styles_(styles), numbering_(numbering) {}

    void scan(pugi::xml_node root, const RelMap& rels) { walk(root, rels); }

    const std::set<std::string>& found() const { return found_; }

private:
    void walk(pugi::xml_node node, const RelMap& rels) {
        for (pugi::xml_node child : node.children()) {
            const std::string_view name = child.name();
            if (name == "mc:Fallback") continue;

            if (name == "w:p") {
                if (pugi::xml_node num_pr = styles_.num_pr(child)) {
                    if (auto fmt = numbering_.unsupported_format(num_pr)) {
                        found_.insert(fmt::format("list numbering format '{}'", *fmt));
                    }
                }
            } else if (name == "w:txbxContent") {
                found_.insert("text boxes");
            } else if (name == "w:footnoteReference" || name == "w:endnoteReference") {
                found_.insert("footnotes and endnotes");
            } else if (name == "w:sym") {
                found_.insert("symbol characters (w:sym)");
            } else if (name == "w:object") {
                found_.insert("embedded objects (w:object)");
            } else if (name == "m:oMath") {
                found_.insert("equations");
            } else if (name == "w:pict") {
                found_.insert("legacy VML graphics (w:pict)");
            } else if (name == "w:cols" && child.attribute("w:num").as_int(1) > 1) {
                found_.insert("multi-column page layout");
            } else if (name == "w:drawing") {
                check_drawing(child, rels);
            }
            walk(child, rels);
        }
    }

    void check_drawing(pugi::xml_node drawing, const RelMap& rels) {
        pugi::xml_node data = find_descendant(drawing, "a:graphicData");
        const std::string uri = data.attribute("uri").value();
        if (uri == kPictureUri) {
            pugi::xml_node blip = find_descendant(data, "a:blip");
            const auto rel = rels.find(blip.attribute("r:embed").value());
            if (rel == rels.end()) return;  // reported as a broken document when drawn
            const auto part = parts_.find("word/" + rel->second);
            if (part == parts_.end()) return;
            const std::string fmt = image_format(part->second);
            if (fmt != "png" && fmt != "jpeg") {
                found_.insert(fmt::format("images in {} format", fmt));
            }
        } else if (uri == kChartUri) {
            const auto rel = rels.find(find_descendant(data, "c:chart").attribute("r:id").value());
            if (rel != rels.end()) check_chart("word/" + rel->second);
        } else if (uri.find("chartex") != std::string::npos) {
            found_.insert("Office 2016+ charts (waterfall, histogram, treemap, ...)");
        } else if (uri.find("diagram") != std::string::npos) {
            found_.insert("SmartArt");
        } else if (uri.find("wordprocessingShape") != std::string::npos ||
                   uri.find("wordprocessingGroup") != std::string::npos ||
                   uri.find("wordprocessingCanvas") != std::string::npos) {
            found_.insert("shapes and drawing groups");
        } else {
            found_.insert(fmt::format("drawings of type '{}'", uri));
        }
    }

    void check_chart(const std::string& part_name) {
        const auto part = parts_.find(part_name);
        if (part == parts_.end()) return;
        pugi::xml_document chart;
        if (!chart.load_buffer(part->second.data(), part->second.size())) return;
        pugi::xml_node plot_area = chart.child("c:chartSpace").child("c:chart").child("c:plotArea");

        std::vector<pugi::xml_node> types;
        for (pugi::xml_node child : plot_area.children()) {
            const std::string_view name = child.name();
            if (name.size() > 7 && name.substr(0, 2) == "c:" &&
                name.substr(name.size() - 5) == "Chart") {
                types.push_back(child);
            }
        }
        bool pie = false, horizontal = false;
        for (pugi::xml_node t : types) {
            const std::string_view name = t.name();
            if (name == "c:ofPieChart")                          found_.insert("pie-of-pie / bar-of-pie charts");
            else if (name == "c:radarChart")                     found_.insert("radar charts");
            else if (name == "c:scatterChart")                   found_.insert("scatter charts");
            else if (name == "c:bubbleChart")                    found_.insert("bubble charts");
            else if (name == "c:stockChart")                     found_.insert("stock charts");
            else if (name == "c:surfaceChart" || name == "c:surface3DChart") found_.insert("surface charts");

            pie = pie || name == "c:pieChart" || name == "c:pie3DChart" || name == "c:doughnutChart";
            if ((name == "c:barChart" || name == "c:bar3DChart") &&
                std::strcmp(t.child("c:barDir").attribute("val").value(), "bar") == 0) {
                horizontal = true;
            }
            if (name == "c:doughnutChart") {
                std::size_t n = 0;
                for (pugi::xml_node ser : t.children("c:ser")) { (void)ser; ++n; }
                if (n > 1) found_.insert("multi-ring doughnut charts");
            }
        }
        // Bar/line/area combinations (stacked or not, secondary axis) are
        // drawn; these mixes are not.
        if (types.size() > 1 && pie)        found_.insert("pie charts combined with other chart types");
        if (types.size() > 1 && horizontal) found_.insert("horizontal bar charts combined with other chart types");
    }

    const PartMap&        parts_;
    const StyleSheet&     styles_;
    const Numbering&      numbering_;
    std::set<std::string> found_;
};

// Explicit env override (mirrors the DOCWEFT_SOFFICE pattern) first, then
// a short list of Unicode TrueType fonts commonly present per platform.
// DejaVu Sans is the priority candidate specifically because it covers
// Cyrillic — DocWeft's own example templates rely on Cyrillic to
// demonstrate UTF-8 support, and Standard 14 PDF fonts (WinAnsi-only) can't
// render it at all.
std::optional<std::string> discover_font_path(bool bold) {
    const char* env = std::getenv(bold ? "DOCWEFT_NATIVE_PDF_FONT_BOLD"
                                        : "DOCWEFT_NATIVE_PDF_FONT");
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

// Gap between a header/footer and the body text.
constexpr double kBlockGap = 6.0;

// ── Fonts ───────────────────────────────────────────────────────────────────
//
// Finds each run's font family by name through PoDoFo's system font search
// (fontconfig / GDI). Only a font of the requested family is accepted — a
// system search may otherwise hand back any installed font (FreeMono for
// "Calibri", seen with a minimal fontconfig setup). A family that isn't
// installed is tried under its metric-compatible free substitute
// (Calibri → Carlito, Arial → Liberation Sans, ...), the ones LibreOffice
// uses, so line breaks stay close to Word's. Otherwise — or when the font
// lacks a glyph the text needs (a Latin-only font and Cyrillic text, say)
// — the fallback font from discover_font_path() is used.
class FontBook {
public:
    FontBook(PoDoFo::PdfMemDocument& doc, std::string default_family)
        : doc_(doc), default_family_(std::move(default_family)) {
        const std::optional<std::string> regular_path = discover_font_path(false);
        const std::optional<std::string> bold_path    = discover_font_path(true);
        fallback_regular_ = regular_path
            ? &doc.GetFonts().GetOrCreateFont(*regular_path)
            : &doc.GetFonts().GetStandard14Font(PoDoFo::PdfStandard14FontType::Helvetica);
        fallback_bold_ = bold_path
            ? &doc.GetFonts().GetOrCreateFont(*bold_path)
            : (regular_path ? fallback_regular_
                            : &doc.GetFonts().GetStandard14Font(PoDoFo::PdfStandard14FontType::HelveticaBold));
    }

    PoDoFo::PdfFont& get(const TextStyle& st, std::string_view text = {}) {
        const std::string& family = st.font.empty() ? default_family_ : st.font;
        if (PoDoFo::PdfFont* f = search(family, st.bold, st.italic); f != nullptr && covers(*f, text)) {
            return *f;
        }
        return st.bold ? *fallback_bold_ : *fallback_regular_;
    }

    // `preferred` if it can draw `text`, else the fallback font.
    PoDoFo::PdfFont& for_text(PoDoFo::PdfFont& preferred, std::string_view text) {
        if (covers(preferred, text)) return preferred;
        return &preferred == fallback_bold_ ? *fallback_bold_ : *fallback_regular_;
    }

private:
    static std::string normalized(std::string_view name) {
        std::string out;
        for (const char ch : name) {
            if (ch != ' ' && ch != '-') out.push_back(static_cast<char>(std::tolower(static_cast<unsigned char>(ch))));
        }
        return out;
    }

    static std::vector<std::string> substitutes(const std::string& family) {
        static const std::map<std::string, std::vector<std::string>> kTable = {
            {"calibri",         {"Carlito"}},
            {"calibrilight",    {"Carlito"}},
            {"cambria",         {"Caladea"}},
            {"arial",           {"Liberation Sans", "Arimo"}},
            {"helvetica",       {"Liberation Sans", "Arimo"}},
            {"arialnarrow",     {"Liberation Sans Narrow"}},
            {"timesnewroman",   {"Liberation Serif", "Tinos"}},
            {"times",           {"Liberation Serif", "Tinos"}},
            {"couriernew",      {"Liberation Mono", "Cousine"}},
            {"courier",         {"Liberation Mono", "Cousine"}},
            {"georgia",         {"Gelasio"}},
        };
        const auto it = kTable.find(normalized(family));
        return it == kTable.end() ? std::vector<std::string>{} : it->second;
    }

    // A font of exactly `family` (by its own family name), or nullptr.
    PoDoFo::PdfFont* find_family(const std::string& family, bool bold, bool italic) {
        PoDoFo::PdfFontSearchParams params;
        params.Style = static_cast<PoDoFo::PdfFontStyle>(
            (bold ? static_cast<unsigned>(PoDoFo::PdfFontStyle::Bold) : 0u) |
            (italic ? static_cast<unsigned>(PoDoFo::PdfFontStyle::Italic) : 0u));
        PoDoFo::PdfFont* font = nullptr;
        try {
            font = doc_.GetFonts().SearchFont(family, params);
        } catch (const PoDoFo::PdfError&) {
            return nullptr;
        }
        if (font == nullptr) return nullptr;
        const std::string found  = normalized(font->GetMetrics().GeFontFamilyNameSafe());
        const std::string wanted = normalized(family);
        // "Calibri Light" may come back as family "Calibri Light" or "Calibri".
        if (found.empty() || wanted.rfind(found, 0) != 0) return nullptr;
        return font;
    }

    PoDoFo::PdfFont* search(const std::string& family, bool bold, bool italic) {
        const std::string key = family + (bold ? "|b" : "|") + (italic ? "i" : "");
        if (const auto it = cache_.find(key); it != cache_.end()) return it->second;
        PoDoFo::PdfFont* font = nullptr;
        if (!family.empty()) {
            font = find_family(family, bold, italic);
            for (const std::string& sub : substitutes(family)) {
                if (font != nullptr) break;
                font = find_family(sub, bold, italic);
            }
        }
        cache_[key] = font;
        return font;
    }

    // Whether `font` has a glyph for every non-space character of `text`.
    static bool covers(PoDoFo::PdfFont& font, std::string_view text) {
        const bool std14 = font.GetMetrics().IsStandard14FontMetrics();
        for (std::size_t i = 0; i < text.size();) {
            const char32_t cp = next_code_point(text, i);
            if (cp <= 0x20) continue;
            if (std14) {
                if (cp > 0xFF) return false;
                continue;
            }
            unsigned gid = 0;
            if (!font.TryGetGID(cp, PoDoFo::PdfGlyphAccess::FontProgram, gid) || gid == 0) return false;
        }
        return true;
    }

    PoDoFo::PdfMemDocument&                  doc_;
    std::string                              default_family_;
    PoDoFo::PdfFont*                         fallback_regular_ = nullptr;
    PoDoFo::PdfFont*                         fallback_bold_    = nullptr;
    std::map<std::string, PoDoFo::PdfFont*>  cache_;
};

// ── Inline layout ───────────────────────────────────────────────────────────

// A measured piece of a line: a word (or part of one in a single style), a
// space, or a tab.
struct Box {
    enum class Kind { Text, Space, Tab } kind = Kind::Text;
    std::string      text;             // Tab: its leader character, if any
    TextStyle        style;
    PoDoFo::PdfFont* font  = nullptr;
    double           width = 0.0;
};

struct Line {
    std::vector<Box> boxes;
    double x0      = 0.0;    // start, from the paragraph's left edge (indents applied)
    double avail   = 0.0;    // width available from x0
    double height  = 0.0;
    double ascent  = 0.0;
    bool   justify = false;  // stretch spaces to fill (justified, not a paragraph's last line)
    bool   page_break_before = false;
};

struct LaidParagraph {
    ParaProps         props;
    std::vector<Line> lines;
    double lines_height() const {
        double h = 0.0;
        for (const Line& l : lines) h += l.height;
        return h;
    }
};

// Spacing between consecutive paragraphs of one container (body, cell,
// header): the previous paragraph's space after plus this one's space
// before, each dropped under contextualSpacing between same-style
// paragraphs (list items, typically).
struct FlowState {
    double      pending_after   = 0.0;
    std::string prev_style;
    bool        prev_contextual = false;
    bool        has_prev        = false;

    double gap_before(const ParaProps& pp) const {
        const bool same = has_prev && prev_style == pp.style_id;
        const double after  = (same && prev_contextual) ? 0.0 : pending_after;
        const double before = (same && pp.contextual) ? 0.0 : pp.before;
        return after + before;
    }
    void finished(const ParaProps& pp) {
        pending_after   = pp.after;
        prev_style      = pp.style_id;
        prev_contextual = pp.contextual;
        has_prev        = true;
    }
    void reset() { *this = FlowState{}; }
};

struct CellMargins {
    double top = 0.0, left = 5.4, bottom = 0.0, right = 5.4;  // Word: 0 / 0.19 cm
};

struct TableLayout;

// One piece of a table cell's content, stacked top to bottom in document
// order, so nothing in a cell is dropped.
// Where a paragraph sits in a bordered/shaded group (consecutive
// paragraphs with ParaProps::same_box()) and the room its box adds above
// and below its lines: border width plus its distance from the text.
struct ParaBox {
    bool   first = true, last = true;  // the group starts / ends here
    double top = 0.0, bottom = 0.0;

    ParaBox() = default;
    ParaBox(const ParaProps& pp, const ParaProps* prev, const ParaProps* next) {
        if (!pp.boxed()) return;
        first = prev == nullptr || !prev->same_box(pp);
        last  = next == nullptr || !pp.same_box(*next);
        const Border& t = first ? pp.borders.top : pp.borders.between;
        if (t.visible()) top = t.space + t.width;
        if (last && pp.borders.bottom.visible()) bottom = pp.borders.bottom.space + pp.borders.bottom.width;
    }
};

struct CellPart {
    enum class Kind { Paragraph, Image, Chart, NestedTable } kind = Kind::Paragraph;
    LaidParagraph   paragraph;
    ImageBlock      image;
    ResolvedChart   chart;
    std::shared_ptr<TableLayout> nested;
    double          y      = 0.0;  // top offset inside the cell's content box
    double          width  = 0.0;  // drawn size (images/charts scaled to the cell)
    double          height = 0.0;
    ParaProps::Align align = ParaProps::Align::Left;  // images/charts: their paragraph's
    // Paragraph box: its top border edge `box_above` over the text, the
    // bottom one `box_below` under it; a group's later paragraph also
    // fills the `box_gap` spacing above it.
    ParaBox         box;
    double          box_above = 0.0, box_below = 0.0, box_gap = 0.0;
};

struct CellContent {
    std::vector<CellPart> parts;
    double                height = 0.0;
};

struct CellLayout {
    pugi::xml_node       tc;
    CellContent          content;
    double               width = 0.0;
    std::size_t          grid_col = 0;     // first <w:tblGrid> column the cell occupies
    std::size_t          rowspan = 1;      // <w:vMerge w:val="restart">: rows it covers
    bool                 covered = false;  // a <w:vMerge/> continuation: drawn by the restart cell
    CellMargins          margins;
    std::optional<Color> fill;
    StyleBorders         style_borders;    // from the table style's conditional formatting
    enum class VAlign { Top, Center, Bottom } valign = VAlign::Top;
};

struct RowLayout {
    std::vector<CellLayout> cells;
    double                  height = 0.0;
    bool                    header = false;  // <w:tblHeader>: repeated on each page
};

struct TableLayout {
    std::vector<RowLayout> rows;
    double                 width    = 0.0;
    double                 x_offset = 0.0;   // from the available area's left edge
    pugi::xml_node         tbl_borders;
    pugi::xml_node         style_borders;
    bool                   style_resolved = false;
    double total_height() const {
        double h = 0.0;
        for (const RowLayout& r : rows) h += r.height;
        return h;
    }
};

// Header or footer content of one kind (default / first page / even pages).
struct HeaderFooterPart {
    pugi::xml_node              root;   // <w:hdr> / <w:ftr>
    std::vector<pugi::xml_node> blocks;
    RelMap                      rels;
    double                      height = 0.0;  // measured before layout
    bool                        present = false;
};

struct HeaderFooterSet {
    enum Kind { Default = 0, First = 1, Even = 2 };
    std::array<HeaderFooterPart, 3> header;
    std::array<HeaderFooterPart, 3> footer;
    bool title_page      = false;  // <w:titlePg>: a section's first page uses First
    bool even_and_odd    = false;  // settings <w:evenAndOddHeaders>: even pages use Even

    Kind kind_for(int page_no, bool first_in_section) const {
        if (title_page && first_in_section) return First;
        if (even_and_odd && page_no % 2 == 0) return Even;
        return Default;
    }
};

// One document section: the blocks up to and including the paragraph whose
// <w:pPr> carries its <w:sectPr> (the last section's sectPr is the body's).
struct Section {
    enum class Start { NextPage, Continuous, EvenPage, OddPage };
    PageMetrics                 pm;
    HeaderFooterSet             hf;
    std::vector<pugi::xml_node> blocks;
    Start                       start = Start::NextPage;
    int                         page_number_start = 0;  // <w:pgNumType w:start>; 0 = continue
    std::string                 page_number_format;     // <w:pgNumType w:fmt>; "" = decimal
};

// Walks one in-memory document and lays it out onto a growing PdfMemDocument,
// page by page. Not reentrant / not thread-safe — one Layout per document.
// What a layout pass found out that fields refer to: the page each
// bookmark landed on (PAGEREF — TOC page numbers are PAGEREF fields) and
// the page counts (NUMPAGES/SECTIONPAGES in body text, which is laid out
// before the count is known).
// A floating picture or chart (<wp:anchor>, outside table cells) placed
// on a page: drawn over the text at the end of its page, or — behindDoc —
// under it, which needs a pass that knows about it before the page's text
// (PoDoFo can only append to a page).
struct FloatingObject {
    pugi::xml_node                 drawing;
    std::size_t                    page = 0;  // index into the pages laid out
    double                         x = 0.0, y = 0.0, w = 0.0, h = 0.0;  // y: bottom edge
    long long                      z = 0;     // <wp:anchor relativeHeight>: stacking order
    std::shared_ptr<ImageBlock>    image;     // one of the two
    std::shared_ptr<ResolvedChart> chart;

    bool operator==(const FloatingObject& o) const {
        return drawing == o.drawing && page == o.page && x == o.x && y == o.y && w == o.w && h == o.h;
    }
};

struct LayoutFacts {
    std::map<std::string, PageNumber> bookmark_pages;  // displayed page numbers
    int                         total = 0;
    std::map<std::size_t, int>  section_pages;
    std::vector<FloatingObject> behind;          // behindDoc floating objects
    bool operator==(const LayoutFacts&) const = default;
};

class Layout {
public:
    // `known`: facts from a previous pass, for PAGEREF/NUMPAGES in the body.
    Layout(PoDoFo::PdfMemDocument& doc, std::vector<Section>& sections, FontBook& fonts,
           const TextStyle& default_style, const pugi::xml_document& styles_doc,
           const StyleSheet& styles, Numbering& numbering,
           const Theme& theme, double default_tab, const LayoutFacts* known = nullptr)
        : doc_(doc), sections_(sections), pm_(sections.front().pm), fonts_(fonts),
          regular_(fonts.get(default_style)),
          bold_(fonts.get([&] { TextStyle b = default_style; b.bold = true; return b; }())),
          styles_doc_(styles_doc), styles_(styles), numbering_(numbering),
          theme_(theme), default_tab_(default_tab > 0.0 ? default_tab : 36.0), known_(known) {}

    LayoutFacts facts() const {
        LayoutFacts f;
        f.bookmark_pages = bookmark_pages_;
        f.behind         = behind_floats_;
        f.total          = static_cast<int>(pages_.size());
        for (const PageInfo& info : pages_) ++f.section_pages[info.section];
        return f;
    }

    // Measures every section's headers/footers once (at that section's
    // text width); page geometry depends on them.
    void measure_headers_footers(const PartMap& parts) {
        for (Section& sec : sections_) {
            pm_ = sec.pm;
            for (auto* set : {&sec.hf.header, &sec.hf.footer}) {
                for (HeaderFooterPart& part : *set) {
                    if (part.present) part.height = measure_blocks(part.blocks, part.rels, parts);
                }
            }
        }
        pm_ = sections_.front().pm;
    }

    // Lays out every section's blocks, starting pages as each section's
    // break type asks.
    void layout_sections(const RelMap& rels, const PartMap& parts) {
        for (std::size_t si = 0; si < sections_.size(); ++si) {
            const Section& sec = sections_[si];
            const Section::Start start = si == 0 ? Section::Start::NextPage : sec.start;
            section_ = si;
            if (start == Section::Start::Continuous) {
                // Same page, new margins for what follows.
                pm_.margin_left_pt  = sec.pm.margin_left_pt;
                pm_.margin_right_pt = sec.pm.margin_right_pt;
            } else {
                pm_ = sec.pm;
                if (sec.page_number_start > 0) next_number_ = sec.page_number_start;
                first_in_section_ = true;
                new_page();
                // Even/odd page breaks insert a blank page when needed.
                const int number = pages_.back().number;
                if ((start == Section::Start::EvenPage && number % 2 != 0) ||
                    (start == Section::Start::OddPage && number % 2 == 0)) {
                    new_page();
                }
            }
            flow_.reset();
            draw_blocks(sec.blocks, rels, parts);
        }
    }

    void new_page() {
        if (page_open_) {
            flush_front_floats();
            painter_.FinishDrawing();
        }
        PoDoFo::PdfPage& page = doc_.GetPages().CreatePage(
            PoDoFo::Rect(0, 0, pm_.width_pt, pm_.height_pt));
        painter_.SetCanvas(page);
        pages_.push_back({section_, next_number_++, first_in_section_});
        first_in_section_ = false;
        page_open_ = true;

        // Body area: below the header / above the footer when those are
        // taller than the page margins leave room for — as Word does.
        const PageInfo& info = pages_.back();
        const HeaderFooterSet& hf = sections_[info.section].hf;
        const auto kind = hf.kind_for(info.number, info.first_in_section);
        const HeaderFooterPart& header = hf.header[kind];
        const HeaderFooterPart& footer = hf.footer[kind];
        page_top_ = pm_.height_pt - pm_.margin_top_pt;
        if (header.present) {
            page_top_ = std::min(page_top_, pm_.height_pt - pm_.header_dist_pt - header.height - kBlockGap);
        }
        bottom_limit_ = pm_.margin_bottom_pt;
        if (footer.present) {
            bottom_limit_ = std::max(bottom_limit_, pm_.footer_dist_pt + footer.height + kBlockGap);
        }
        cursor_y_ = page_top_;
        float_page_ = pages_.size() - 1;
        bands_.clear();
        draw_known_behind(float_page_);
    }

    void finish() {
        if (page_open_) {
            flush_front_floats();
            painter_.FinishDrawing();
            page_open_ = false;
        }
    }

    // Draws headers and footers onto every page once the body is laid out,
    // so NUMPAGES is known. Call after finish().
    void draw_headers_footers(const PartMap& parts) {
        std::map<std::size_t, int> section_pages;
        for (const PageInfo& info : pages_) ++section_pages[info.section];
        for (std::size_t i = 0; i < pages_.size(); ++i) {
            const PageInfo& info = pages_[i];
            const Section& sec = sections_[info.section];
            const auto kind = sec.hf.kind_for(info.number, info.first_in_section);
            const HeaderFooterPart& header = sec.hf.header[kind];
            const HeaderFooterPart& footer = sec.hf.footer[kind];
            if (!header.present && !footer.present) continue;

            painter_.SetCanvas(doc_.GetPages().GetPageAt(static_cast<unsigned>(i)));
            float_page_    = i;
            pm_            = sec.pm;
            fields_        = {info.number, static_cast<int>(pages_.size()), section_pages[info.section],
                              sec.page_number_format, known_ != nullptr ? &known_->bookmark_pages : nullptr};
            no_page_break_ = true;
            if (header.present) {
                cursor_y_ = pm_.height_pt - pm_.header_dist_pt;
                flow_.reset();
                draw_blocks(header.blocks, header.rels, parts);
            }
            if (footer.present) {
                cursor_y_ = pm_.footer_dist_pt + footer.height;
                flow_.reset();
                draw_blocks(footer.blocks, footer.rels, parts);
            }
            no_page_break_ = false;
            flush_front_floats();
            painter_.FinishDrawing();
        }
    }

    void draw_blocks(const std::vector<pugi::xml_node>& blocks, const RelMap& rels,
                     const PartMap& parts) {
        std::vector<pugi::xml_node> shown;
        for (pugi::xml_node b : blocks) {
            if (!hidden_paragraph(b)) shown.push_back(b);
        }
        for (std::size_t i = 0; i < shown.size(); ++i) {
            const pugi::xml_node b = shown[i];
            if (std::strcmp(b.name(), "w:tbl") == 0) {
                draw_table(b, rels, parts);
            } else {
                draw_paragraph(b, rels, parts, i > 0 ? shown[i - 1] : pugi::xml_node{},
                               i + 1 < shown.size() ? shown[i + 1] : pugi::xml_node{});
            }
        }
    }

    void draw_paragraph(pugi::xml_node p, const RelMap& rels, const PartMap& parts,
                        pugi::xml_node prev = {}, pugi::xml_node next = {}) {
        if (on_off(p.child("w:pPr").child("w:pageBreakBefore"))) break_page();

        const LaidParagraph lp = layout_paragraph(p, content_width());
        const ParaProps& pp = lp.props;
        const std::optional<ParaProps> prev_pp = props_of(prev), next_pp = props_of(next);
        const ParaBox box(pp, prev_pp ? &*prev_pp : nullptr, next_pp ? &*next_pp : nullptr);
        const double gap = flow_.gap_before(lp.props);

        // <w:keepNext> (headings, captions): move to the next page rather
        // than stay at the bottom of this one without what follows.
        if (lp.props.keep_next && next && !dry_run_ && !no_page_break_ && cursor_y_ < page_top_) {
            double own = gap + box.top + lp.lines_height() + box.bottom;
            for (pugi::xml_node drawing : flow_drawings(p)) own += drawing_height(drawing, rels, parts);
            const double together = own + first_part_height(next, rels, parts);
            if (cursor_y_ - together < bottom_limit_ && together < page_top_ - bottom_limit_) new_page();
        }
        // A box continued from the previous paragraph also covers the gap.
        const std::size_t pages_before = pages_.size();
        double seg_top = cursor_y_;
        advance(gap);
        if (box.first || pages_.size() != pages_before) seg_top = cursor_y_;
        std::optional<double> top_line = cursor_y_;  // the box's top/between border, until drawn
        const double area_left = pm_.margin_left_pt, area_width = content_width();
        fill_box(pp, area_left, area_width, seg_top, cursor_y_ - box.top);
        cursor_y_ -= box.top;

        // Floating pictures/charts are placed relative to where the
        // paragraph starts: here, or at its first line once that has found
        // its page.
        bool floats_placed = false;
        auto place = [&] {
            if (floats_placed) return;
            floats_placed = true;
            place_floats(p, rels, parts, lp.props, cursor_y_);
        };
        if (lp.lines.empty() || !flow_drawings(p).empty()) place();

        // Pictures and charts first, then the paragraph's own text (if any):
        // a paragraph holding both keeps both, just not interleaved.
        for (pugi::xml_node drawing : flow_drawings(p)) {
            if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
                draw_chart_block(drawing, chart_ref, rels, parts, lp.props);
            } else {
                draw_image(resolve_image(drawing, rels, parts), lp.props);
            }
        }
        for (const Line& line : lp.lines) {
            // A box broken by a page: its lines so far are finished on this page.
            const bool breaks = !dry_run_ && !no_page_break_ && cursor_y_ < page_top_ &&
                                (line.page_break_before || cursor_y_ - line.height < bottom_limit_);
            if (breaks && pp.boxed()) {
                stroke_box(pp, box, area_left, area_width, seg_top, cursor_y_, top_line, false);
                top_line.reset();
            }
            if (line.page_break_before) break_page();
            ensure_space(line.height);
            if (&line == &lp.lines.front() && !floats_placed) {
                place();
                ensure_space(line.height);  // its own floating objects may push it down
            }
            if (breaks) seg_top = cursor_y_;
            if (&line == &lp.lines.front()) record_bookmarks(p);
            // Shading strip by strip, each under its own text.
            fill_box(pp, area_left, area_width, cursor_y_, cursor_y_ - line.height);
            draw_text_line(line, lp.props, pm_.margin_left_pt, cursor_y_);
            cursor_y_ -= line.height;
        }
        if (lp.lines.empty()) record_bookmarks(p);
        fill_box(pp, area_left, area_width, cursor_y_, cursor_y_ - box.bottom);
        cursor_y_ -= box.bottom;
        if (pp.boxed()) stroke_box(pp, box, area_left, area_width, seg_top, cursor_y_, top_line, true);
        flow_.finished(lp.props);
    }

    // ── Floating pictures and charts (<wp:anchor>) ─────────────────────────
    //
    // Placed on the page their paragraph starts on, where <wp:positionH>/
    // <wp:positionV> say: an offset or an alignment within the page, the
    // margins, the paragraph/line, the character position (taken as the
    // paragraph's indent). Wrapping: wrapNone takes no room; every other
    // kind reserves the picture's band across the page and text continues
    // under it (wrapSquare/Tight/Through flow text beside it in Word — not
    // done here, the side stays empty). In front of the text: drawn when
    // the page is finished; behindDoc: drawn first on the page, which
    // takes a second layout pass (see LayoutFacts).

    void place_floats(pugi::xml_node p, const RelMap& rels, const PartMap& parts, const ParaProps& pp,
                      double para_top) {
        if (dry_run_ || in_cell_ || pages_.empty()) return;
        for (pugi::xml_node drawing : visible_drawings(p)) {
            pugi::xml_node anchor = drawing.child("wp:anchor");
            if (!anchor) continue;
            FloatingObject f;
            f.drawing = drawing;
            f.page    = float_page_;
            f.z       = anchor.attribute("relativeHeight").as_llong(0);
            if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
                f.chart = std::make_shared<ResolvedChart>(resolve_chart(drawing, chart_ref, rels, parts, theme_));
                f.w = f.chart->width_pt;
                f.h = f.chart->height_pt;
            } else {
                f.image = std::make_shared<ImageBlock>(resolve_image(drawing, rels, parts));
                f.w = f.image->width_pt;
                f.h = f.image->height_pt;
            }
            position_float(anchor, pp, para_top, f);

            const bool behind = on_off_attr(anchor.attribute("behindDoc"));
            if (!anchor.child("wp:wrapNone") && !behind) {
                const double dist_t = emu_to_pt(anchor.attribute("distT").as_double(0.0));
                const double dist_b = emu_to_pt(anchor.attribute("distB").as_double(0.0));
                bands_.push_back({f.y + f.h + dist_t, f.y - dist_b});
            }
            if (!behind) {
                front_floats_.push_back(std::move(f));
                continue;
            }
            // Already drawn under the page's text if the previous pass
            // placed it the same; otherwise drawn now (over what's there
            // already) and the next pass gets it right.
            const bool known = known_ != nullptr &&
                std::find(known_->behind.begin(), known_->behind.end(), f) != known_->behind.end();
            if (!known) draw_float(f);
            behind_floats_.push_back(std::move(f));
        }
    }

    static bool on_off_attr(pugi::xml_attribute a) {
        const std::string_view v = a.value();
        return v == "1" || v == "true" || v == "on";
    }

    // Sets f.x/f.y (bottom-left, points) from the anchor's position.
    void position_float(pugi::xml_node anchor, const ParaProps& pp, double para_top, FloatingObject& f) const {
        const double page_w = pm_.width_pt, page_h = pm_.height_pt;
        auto offset = [](pugi::xml_node pos) { return emu_to_pt(pos.child("wp:posOffset").text().as_double(0.0)); };

        if (on_off_attr(anchor.attribute("simplePos"))) {
            pugi::xml_node sp = anchor.child("wp:simplePos");
            f.x = emu_to_pt(sp.attribute("x").as_double(0.0));
            f.y = page_h - emu_to_pt(sp.attribute("y").as_double(0.0)) - f.h;
            return;
        }

        // Horizontal frame: left edge and width.
        pugi::xml_node ph = anchor.child("wp:positionH");
        const std::string_view hrel = ph.attribute("relativeFrom").value();
        double fx = pm_.margin_left_pt, fw = content_width();  // margin, column
        if (hrel == "page") {
            fx = 0.0;
            fw = page_w;
        } else if (hrel == "leftMargin" || hrel == "insideMargin") {
            fx = 0.0;
            fw = pm_.margin_left_pt;
        } else if (hrel == "rightMargin" || hrel == "outsideMargin") {
            fx = page_w - pm_.margin_right_pt;
            fw = pm_.margin_right_pt;
        } else if (hrel == "character") {
            fx = pm_.margin_left_pt + pp.ind_left;
            fw = std::max(content_width() - pp.ind_left, 0.0);
        }
        if (pugi::xml_node a = ph.child("wp:align")) {
            const std::string_view v = a.text().get();
            f.x = v == "center"                     ? fx + (fw - f.w) / 2.0
                : v == "right" || v == "outside"    ? fx + fw - f.w
                                                    : fx;
        } else {
            f.x = fx + offset(ph);
        }

        // Vertical frame: top edge (PDF y) and height.
        pugi::xml_node pv = anchor.child("wp:positionV");
        const std::string_view vrel = pv.attribute("relativeFrom").value();
        double ft = para_top, fh = 0.0;  // paragraph, line
        if (vrel == "page") {
            ft = page_h;
            fh = page_h;
        } else if (vrel == "margin") {
            ft = page_h - pm_.margin_top_pt;
            fh = page_h - pm_.margin_top_pt - pm_.margin_bottom_pt;
        } else if (vrel == "topMargin" || vrel == "insideMargin") {
            ft = page_h;
            fh = pm_.margin_top_pt;
        } else if (vrel == "bottomMargin" || vrel == "outsideMargin") {
            ft = pm_.margin_bottom_pt;
            fh = pm_.margin_bottom_pt;
        }
        double top = ft - offset(pv);
        if (pugi::xml_node a = pv.child("wp:align")) {
            const std::string_view v = a.text().get();
            top = v == "center" ? ft - (fh - f.h) / 2.0 : v == "bottom" ? ft - fh + f.h : ft;
        }
        f.y = top - f.h;
    }

    void draw_float(const FloatingObject& f) {
        if (dry_run_) return;
        if (f.chart) {
            draw_chart_in_box(*f.chart, f.x, f.y, f.w, f.h);
            return;
        }
        std::unique_ptr<PoDoFo::PdfImage> image = doc_.CreateImage();
        image->LoadFromBuffer(PoDoFo::bufferview(f.image->png_bytes.data(), f.image->png_bytes.size()));
        painter_.DrawImage(*image, f.x, f.y, f.w / static_cast<double>(image->GetWidth()),
                           f.h / static_cast<double>(image->GetHeight()));
    }

    // In-front floating objects of the page being finished, bottom to top.
    void flush_front_floats() {
        std::stable_sort(front_floats_.begin(), front_floats_.end(),
                         [](const FloatingObject& a, const FloatingObject& b) { return a.z < b.z; });
        for (const FloatingObject& f : front_floats_) draw_float(f);
        front_floats_.clear();
    }

    // behindDoc objects the previous pass placed on page `page`, drawn
    // before anything else on it.
    void draw_known_behind(std::size_t page) {
        if (known_ == nullptr) return;
        std::vector<const FloatingObject*> here;
        for (const FloatingObject& f : known_->behind) {
            if (f.page == page) here.push_back(&f);
        }
        std::stable_sort(here.begin(), here.end(),
                         [](const FloatingObject* a, const FloatingObject* b) { return a->z < b->z; });
        for (const FloatingObject* f : here) draw_float(*f);
    }

    // ── Paragraph borders and shading ───────────────────────────────────────

    // Formatting of `block` if it's a paragraph (list counters untouched).
    std::optional<ParaProps> props_of(pugi::xml_node block) const {
        if (!block || std::strcmp(block.name(), "w:p") != 0) return std::nullopt;
        const pugi::xml_node num_pr = styles_.num_pr(block);
        return styles_.para_props(block, num_pr ? numbering_.level_of(num_pr) : pugi::xml_node{}, cell_style_);
    }

    // Left and right outer edges of a paragraph's box in a text area: the
    // indents, pushed out by the left/right border and its distance.
    static std::pair<double, double> box_edges(const ParaProps& pp, double area_left, double area_width) {
        const Border& l = pp.borders.left;
        const Border& r = pp.borders.right;
        double left  = area_left + std::min(pp.ind_left, pp.ind_left + pp.ind_first);
        double right = area_left + area_width - pp.ind_right;
        if (l.visible()) left -= l.space + l.width;
        if (r.visible()) right += r.space + r.width;
        return {left, right};
    }

    void fill_box(const ParaProps& pp, double area_left, double area_width, double top, double bottom) {
        if (dry_run_ || !pp.shading || top <= bottom) return;
        const auto [left, right] = box_edges(pp, area_left, area_width);
        painter_.GraphicsState.SetNonStrokingColor(pdf_color(*pp.shading));
        painter_.DrawRectangle(left, bottom, right - left, top - bottom, PoDoFo::PdfPathDrawMode::Fill);
        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
    }

    // Border lines of a box's part from `top` down to `bottom` (outer
    // edges): left and right always; the top (group start) or between
    // border at `top_line` if given; the bottom one if `bottom_edge` and
    // the group ends here.
    void stroke_box(const ParaProps& pp, const ParaBox& box, double area_left, double area_width,
                    double top, double bottom, std::optional<double> top_line, bool bottom_edge) {
        if (dry_run_) return;
        const auto [left, right] = box_edges(pp, area_left, area_width);
        const ParaProps::Borders& b = pp.borders;
        if (top_line) {
            const Border& t = box.first ? b.top : b.between;
            const double y = *top_line - t.width / 2.0;
            stroke_border(t, left, y, right, y);
        }
        if (bottom_edge && box.last) {
            const double y = bottom + b.bottom.width / 2.0;
            stroke_border(b.bottom, left, y, right, y);
        }
        stroke_border(b.left, left + b.left.width / 2.0, bottom, left + b.left.width / 2.0, top);
        stroke_border(b.right, right - b.right.width / 2.0, bottom, right - b.right.width / 2.0, top);
    }

    // Top-level table: the only one that participates in page-break
    // decisions (checked per row). Header rows (<w:tblHeader>) are repeated
    // at the top of every page the table continues on. A nested table
    // (reached through a cell's own content) never breaks pages.
    void draw_table(pugi::xml_node tbl, const RelMap& rels, const PartMap& parts) {
        const TableLayout t = measure_table(tbl, content_width(), /*force_fit*/false, rels, parts);
        if (t.rows.empty()) return;
        advance(flow_.pending_after);
        flow_.reset();

        std::size_t header_rows = 0;
        while (header_rows < t.rows.size() && t.rows[header_rows].header) ++header_rows;
        double header_height = 0.0;
        for (std::size_t i = 0; i < header_rows; ++i) header_height += t.rows[i].height;

        const double x = pm_.margin_left_pt + t.x_offset;
        std::size_t keep_until = 0;  // rows before this one were checked as part of a merged block
        for (std::size_t ri = 0; ri < t.rows.size(); ++ri) {
            const RowLayout& row = t.rows[ri];
            double block_height = row.height;
            if (ri >= keep_until) {
                keep_until = merged_block_end(t, ri);
                block_height = 0.0;
                for (std::size_t k = ri; k < keep_until; ++k) block_height += t.rows[k].height;
            } else {
                block_height = 0.0;  // part of a block that already fitted
            }
            if (block_height > 0.0 && !dry_run_ && !no_page_break_) avoid_bands(block_height);
            if (block_height > 0.0 && !dry_run_ && !no_page_break_ && cursor_y_ - block_height < bottom_limit_ &&
                cursor_y_ < page_top_) {
                new_page();
                // Repeat header rows, unless they'd fill the page themselves.
                if (ri >= header_rows && header_rows > 0 &&
                    header_height + block_height < page_top_ - bottom_limit_) {
                    for (std::size_t hi = 0; hi < header_rows; ++hi) {
                        draw_row(t, hi, x, cursor_y_);
                        cursor_y_ -= t.rows[hi].height;
                    }
                }
            }
            draw_row(t, ri, x, cursor_y_);
            cursor_y_ -= row.height;
        }
    }

private:
    // ── Look-ahead for keepNext ─────────────────────────────────────────────

    // Height a picture/chart paragraph's drawing takes in the page flow.
    double drawing_height(pugi::xml_node drawing, const RelMap& rels, const PartMap& parts) {
        constexpr double kLegendRowHeight = 16.0;
        if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
            return resolve_chart(drawing, chart_ref, rels, parts, theme_).height_pt + kLegendRowHeight;
        }
        const ImageBlock img = resolve_image(drawing, rels, parts);
        const double w = img.width_pt;
        return w > content_width() && w > 0.0 ? img.height_pt * content_width() / w : img.height_pt;
    }

    // Height of the first thing `block` puts on a page: a table's first
    // row, or a paragraph's pictures and first line. List counters are
    // left untouched.
    double first_part_height(pugi::xml_node block, const RelMap& rels, const PartMap& parts) {
        const Numbering::Snapshot saved = numbering_.snapshot();
        double h = 0.0;
        if (std::strcmp(block.name(), "w:tbl") == 0) {
            const TableLayout t = measure_table(block, content_width(), false, rels, parts);
            if (!t.rows.empty()) {
                for (std::size_t k = 0; k < merged_block_end(t, 0); ++k) h += t.rows[k].height;
            }
        } else {
            const LaidParagraph lp = layout_paragraph(block, content_width());
            h = lp.props.before;
            for (pugi::xml_node drawing : flow_drawings(block)) h += drawing_height(drawing, rels, parts);
            if (!lp.lines.empty()) h += lp.lines.front().height;
        }
        numbering_.restore(saved);
        return h;
    }

    // ── Paragraph layout ────────────────────────────────────────────────────

    struct Metrics {
        double line   = 0.0;  // single line height
        double ascent = 0.0;
        double descent = 0.0;
    };

    Metrics metrics_of(const TextStyle& st) {
        PoDoFo::PdfFont& font = fonts_.get(st);
        PoDoFo::PdfTextState state;
        state.Font     = &font;
        state.FontSize = st.drawn_size();
        Metrics m;
        m.ascent  = font.GetAscent(state);
        m.descent = std::abs(font.GetDescent(state));
        m.line    = std::max(font.GetLineSpacing(state), m.ascent + m.descent);
        return m;
    }

    double text_width(PoDoFo::PdfFont& font, double size, std::string_view text) {
        PoDoFo::PdfTextState state;
        state.Font     = &font;
        state.FontSize = size;
        return font.GetStringLength(text, state);
    }

    Box make_box(Box::Kind kind, std::string text, const TextStyle& st) {
        Box b;
        b.kind  = kind;
        b.style = st;
        b.font  = &fonts_.get(st, kind == Box::Kind::Text ? std::string_view(text) : std::string_view(" "));
        b.text  = std::move(text);
        if (kind == Box::Kind::Text)       b.width = text_width(*b.font, st.drawn_size(), b.text);
        else if (kind == Box::Kind::Space) b.width = text_width(*b.font, st.drawn_size(), " ");
        return b;
    }

    // Lays `p` out at `width` points: numbering marker, styled words broken
    // into lines (greedy), tabs, line and page breaks, line heights.
    LaidParagraph layout_paragraph(pugi::xml_node p, double width) {
        LaidParagraph lp;
        const pugi::xml_node num_pr = styles_.num_pr(p);
        const pugi::xml_node level  = num_pr ? numbering_.level_of(num_pr) : pugi::xml_node{};
        lp.props = styles_.para_props(p, level, cell_style_);
        const ParaProps& pp = lp.props;
        const TextStyle mark = styles_.mark_style(p, cell_style_);

        std::optional<Numbering::Marker> marker;
        if (num_pr) marker = numbering_.next(num_pr);

        const std::vector<Fragment> fragments = fragments_of(p);
        if (fragments.empty() && !flow_drawings(p).empty() && !(marker && !marker->text.empty())) {
            return lp;  // a paragraph that only holds pictures/charts
        }

        Line cur;
        double x = 0.0;                 // pen position from cur.x0
        std::vector<Box> pending;       // spaces waiting for the next word

        // Right/center tab waiting for the text after it to be measured.
        struct PendingTab { std::size_t box = 0; double stop = 0.0; double start_x = 0.0;
                            ParaProps::Tab::Kind kind = ParaProps::Tab::Kind::Left; bool active = false; } ptab;

        auto start_line = [&](bool first) {
            cur = Line{};
            cur.x0    = pp.ind_left + (first ? pp.ind_first : 0.0);
            cur.avail = std::max(width - cur.x0 - pp.ind_right, 10.0);
            x = 0.0;
            pending.clear();
            ptab.active = false;
        };
        auto settle_tab = [&] {
            if (!ptab.active) return;
            const double seg = x - ptab.start_x;
            const double want = ptab.kind == ParaProps::Tab::Kind::Right ? ptab.stop - seg : ptab.stop - seg / 2.0;
            const double w = std::max(0.0, want - (cur.x0 + ptab.start_x));
            cur.boxes[ptab.box].width = w;
            x += w;
            ptab.active = false;
        };
        auto end_line = [&](bool justify) {
            settle_tab();
            while (!cur.boxes.empty() && cur.boxes.back().kind == Box::Kind::Space) cur.boxes.pop_back();
            Metrics m;
            double unshifted_ascent = 0.0, unshifted_descent = 0.0;
            bool any = false;
            for (const Box& b : cur.boxes) {
                const Metrics bm = metrics_of(b.style);
                const double rise = b.style.rise();
                m.line = std::max(m.line, bm.line);
                m.ascent = std::max(m.ascent, bm.ascent + rise);
                m.descent = std::max(m.descent, bm.descent - rise);
                unshifted_ascent = std::max(unshifted_ascent, bm.ascent);
                unshifted_descent = std::max(unshifted_descent, bm.descent);
                any = true;
            }
            if (!any) m = metrics_of(mark);
            // Raised or lowered text that sticks out makes the line taller,
            // keeping its usual leading.
            if (any) {
                const double leading = std::max(0.0, m.line - unshifted_ascent - unshifted_descent);
                m.line = std::max(m.line, m.ascent + m.descent + leading);
            }
            switch (pp.line_rule) {
                case ParaProps::LineRule::Auto:    cur.height = m.line * pp.line; break;
                case ParaProps::LineRule::Exact:   cur.height = pp.line; break;
                case ParaProps::LineRule::AtLeast: cur.height = std::max(m.line, pp.line); break;
            }
            cur.ascent = pp.line_rule == ParaProps::LineRule::Exact
                ? std::max(0.0, cur.height - m.descent) : m.ascent + std::max(0.0, m.line - m.ascent - m.descent) / 2.0;
            cur.justify = justify && pp.align == ParaProps::Align::Justify;
            lp.lines.push_back(std::move(cur));
            start_line(false);
        };

        start_line(true);

        // List marker, then a tab to the text indent (Word's default suffix).
        if (marker && !marker->text.empty()) {
            TextStyle mst = mark;
            if (!fragments.empty()) {
                mst = fragments.front().style;
                if (fragments.front().small_caps_lower) mst.size_pt /= kSmallCapsScale;
                mst.vert_align  = TextStyle::VertAlign::Baseline;
                mst.position_pt = 0.0;
            }
            const std::string font_before = mst.font;
            styles_.apply_rpr(marker->level.child("w:rPr"), mst);
            mst.font = font_before;  // Symbol/Wingdings: bullets are mapped to Unicode instead
            mst.underline = false;
            Box mb = make_box(Box::Kind::Text, marker->text, mst);
            x += mb.width;
            cur.boxes.push_back(std::move(mb));
            const double text_start = -std::min(pp.ind_first, 0.0);  // hanging → the text indent
            Box gap = make_box(Box::Kind::Tab, "", mst);
            gap.width = text_start > x + 1.0 ? text_start - x
                                             : std::max(default_tab_ - std::fmod(cur.x0 + x, default_tab_), mst.size_pt * 0.4);
            x += gap.width;
            cur.boxes.push_back(std::move(gap));
        }

        auto place_word = [&](std::vector<Box>& word) {
            double w = 0.0;
            for (const Box& b : word) w += b.width;
            double spaces = 0.0;
            for (const Box& b : pending) spaces += b.width;
            const bool has_content = std::any_of(cur.boxes.begin(), cur.boxes.end(),
                                                 [](const Box& b) { return b.kind != Box::Kind::Space; });
            if (has_content && !ptab.active && x + spaces + w > cur.avail + 0.01) {
                end_line(true);  // the spaces before the word end the line
            } else {
                for (Box& b : pending) { x += b.width; cur.boxes.push_back(std::move(b)); }
            }
            pending.clear();

            if (x + w <= cur.avail + 0.01 || ptab.active) {
                for (Box& b : word) { x += b.width; cur.boxes.push_back(std::move(b)); }
                word.clear();
                return;
            }
            // Longer than the rest of an empty line: break it between characters.
            for (Box& piece : word) {
                std::string chunk;
                for (std::size_t i = 0; i < piece.text.size();) {
                    const auto b0 = static_cast<unsigned char>(piece.text[i]);
                    const std::size_t len = b0 >= 0xF0 ? 4 : b0 >= 0xE0 ? 3 : b0 >= 0xC0 ? 2 : 1;
                    const std::string ch = piece.text.substr(i, len);
                    i += len;
                    const double chunk_w = text_width(*piece.font, piece.style.drawn_size(), chunk + ch);
                    if (x + chunk_w > cur.avail && (x > 0.0 || !chunk.empty())) {
                        if (!chunk.empty()) {
                            Box part = make_box(Box::Kind::Text, chunk, piece.style);
                            x += part.width;
                            cur.boxes.push_back(std::move(part));
                            chunk.clear();
                        }
                        end_line(true);
                    }
                    chunk += ch;
                }
                if (!chunk.empty()) {
                    Box part = make_box(Box::Kind::Text, chunk, piece.style);
                    x += part.width;
                    cur.boxes.push_back(std::move(part));
                }
            }
            word.clear();
        };

        std::vector<Box> word;
        auto flush_piece = [&](std::string& text, const TextStyle& st) {
            if (text.empty()) return;
            word.push_back(make_box(Box::Kind::Text, text, st));
            text.clear();
        };

        for (const Fragment& f : fragments) {
            switch (f.kind) {
                case Fragment::Kind::Text: {
                    std::string piece;
                    for (const char ch : f.text) {
                        if (ch == ' ') {
                            flush_piece(piece, f.style);
                            if (!word.empty()) place_word(word);
                            pending.push_back(make_box(Box::Kind::Space, " ", f.style));
                        } else if (static_cast<unsigned char>(ch) < 0x20) {
                            piece += ' ';  // stray control characters
                        } else {
                            piece += ch;
                        }
                    }
                    flush_piece(piece, f.style);
                    break;
                }
                case Fragment::Kind::Tab: {
                    if (!word.empty()) place_word(word);
                    for (Box& b : pending) { x += b.width; cur.boxes.push_back(std::move(b)); }
                    pending.clear();
                    settle_tab();
                    const double abs = cur.x0 + x;
                    ParaProps::Tab stop{ParaProps::Tab::Kind::Left, 0.0};
                    bool found = false;
                    // A hanging indent acts as an implicit tab stop.
                    if (pp.ind_first < 0.0 && abs < pp.ind_left - 0.5) { stop.pos = pp.ind_left; found = true; }
                    for (const ParaProps::Tab& t : pp.tabs) {
                        if (!found && t.pos > abs + 0.5) { stop = t; found = true; }
                    }
                    if (!found) stop.pos = (std::floor(abs / default_tab_ + 1e-6) + 1.0) * default_tab_;
                    Box tab = make_box(Box::Kind::Tab, "", f.style);
                    tab.text = stop.leader;
                    if (stop.kind == ParaProps::Tab::Kind::Left) {
                        tab.width = std::max(0.0, stop.pos - abs);
                        x += tab.width;
                        cur.boxes.push_back(std::move(tab));
                    } else {
                        cur.boxes.push_back(std::move(tab));
                        ptab = {cur.boxes.size() - 1, stop.pos, x, stop.kind, true};
                    }
                    break;
                }
                case Fragment::Kind::LineBreak:
                    if (!word.empty()) place_word(word);
                    end_line(false);
                    break;
                case Fragment::Kind::PageBreak:
                    // The text after the break starts the next page; a break
                    // at the very start of a paragraph moves the whole of it.
                    if (!word.empty()) place_word(word);
                    if (!cur.boxes.empty() || !lp.lines.empty()) end_line(false);
                    cur.page_break_before = true;
                    break;
            }
        }
        if (!word.empty()) place_word(word);
        end_line(false);  // also the empty line a trailing page break leaves, as in Word
        return lp;
    }

    // Draws one laid-out line whose top is at `top`; `left` is the
    // paragraph's left edge.
    void draw_text_line(const Line& line, const ParaProps& pp, double left, double top) {
        if (dry_run_) return;
        double natural = 0.0;
        std::size_t spaces = 0;
        for (const Box& b : line.boxes) {
            natural += b.width;
            if (b.kind == Box::Kind::Space) ++spaces;
        }
        const double extra = std::max(0.0, line.avail - natural);
        double x = left + line.x0;
        if (pp.align == ParaProps::Align::Center)     x += extra / 2.0;
        else if (pp.align == ParaProps::Align::Right) x += extra;
        const double stretch = (line.justify && spaces > 0) ? extra / static_cast<double>(spaces) : 0.0;
        const double baseline = top - line.ascent;

        // Consecutive boxes in the same style are drawn with one text call;
        // a justified line draws word by word, stretching the spaces.
        std::size_t i = 0;
        while (i < line.boxes.size()) {
            const Box& b = line.boxes[i];
            if (b.kind == Box::Kind::Tab) {
                draw_leader(b, x, baseline + b.style.rise());
                x += b.width;
                ++i;
                continue;
            }
            if (stretch > 0.0) {
                if (b.kind == Box::Kind::Space) {
                    x += b.width + stretch;
                } else {
                    draw_styled(*b.font, b.style, b.text, x, baseline + b.style.rise(), b.width);
                    x += b.width;
                }
                ++i;
                continue;
            }
            std::string text;
            double width = 0.0;
            bool has_glyphs = false;
            std::size_t j = i;
            while (j < line.boxes.size() && line.boxes[j].kind != Box::Kind::Tab &&
                   line.boxes[j].style == b.style && line.boxes[j].font == b.font) {
                const Box& g = line.boxes[j];
                text  += g.kind == Box::Kind::Space ? std::string(" ") : g.text;
                width += g.width;
                has_glyphs = has_glyphs || g.kind == Box::Kind::Text;
                ++j;
            }
            if (has_glyphs || b.style.underline || b.style.strike) {
                draw_styled(*b.font, b.style, text, x, baseline + b.style.rise(), width);
            }
            x += width;
            i = j;
        }
    }

    // A tab's leader (dots, hyphens, ...) across its width. The characters
    // sit on a grid of their own width from the page's left edge, so
    // leaders on consecutive lines (a table of contents) line up; a little
    // room is kept clear of the text on either side.
    void draw_leader(const Box& tab, double x, double baseline) {
        if (tab.text.empty() || tab.width <= 0.0) return;
        TextStyle st = tab.style;
        st.underline = st.strike = false;
        PoDoFo::PdfFont& font = fonts_.get(st, tab.text);
        const double w = text_width(font, st.drawn_size(), tab.text);
        if (w <= 0.0) return;
        const double clear = st.drawn_size() * 0.15;
        const double start = std::ceil((x + clear) / w - 1e-9) * w;
        std::string run;
        double end = start;
        while (end + w <= x + tab.width - clear + 1e-6) {
            run += tab.text;
            end += w;
        }
        if (!run.empty()) draw_styled(font, st, run, start, baseline, end - start);
    }

    void draw_styled(PoDoFo::PdfFont& font, const TextStyle& st, const std::string& text,
                     double x, double baseline, double width) {
        painter_.GraphicsState.SetNonStrokingColor(pdf_color(st.color));
        painter_.TextState.SetFont(font, st.drawn_size());
        painter_.DrawText(sanitize_single_line(text), x, baseline);
        if (st.underline || st.strike) {
            PoDoFo::PdfTextState state;
            state.Font     = &font;
            state.FontSize = st.drawn_size();
            painter_.GraphicsState.SetStrokingColor(pdf_color(st.color));
            if (st.underline) {
                painter_.GraphicsState.SetLineWidth(std::max(font.GetUnderlineThickness(state), 0.5));
                const double y = baseline + font.GetUnderlinePosition(state);
                painter_.DrawLine(x, y, x + width, y);
            }
            if (st.strike) {
                painter_.GraphicsState.SetLineWidth(std::max(font.GetStrikeThroughThickness(state), 0.5));
                const double y = baseline + font.GetStrikeThroughPosition(state);
                painter_.DrawLine(x, y, x + width, y);
            }
            painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            painter_.GraphicsState.SetLineWidth(1.0);
        }
        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
    }

    // ── Tables ──────────────────────────────────────────────────────────────

    // Cell margins: the table's <w:tblCellMar> (own, else its style's),
    // then the cell's <w:tcMar>.
    static void apply_margins(pugi::xml_node mar, CellMargins& m) {
        if (!mar) return;
        auto side = [&](const char* a, const char* b, double& out) {
            for (const char* name : {a, b}) {
                if (name == nullptr) continue;
                if (pugi::xml_node n = mar.child(name)) {
                    if (std::strcmp(n.attribute("w:type").value(), "pct") != 0) {
                        out = twips_to_pt(n.attribute("w:w").as_llong(0));
                    }
                    return;
                }
            }
        };
        side("w:top", nullptr, m.top);
        side("w:left", "w:start", m.left);
        side("w:bottom", nullptr, m.bottom);
        side("w:right", "w:end", m.right);
    }

    // ── Table style conditional formatting ─────────────────────────────────
    //
    // A table style formats regions of the table (<w:tblStylePr>): header
    // and total rows, first and last columns, banded rows and columns,
    // corner cells — text (rPr), paragraphs (pPr), shading and borders
    // (tcPr). <w:tblLook> says which regions are on. Later regions win, in
    // the order LibreOffice applies them (checked on a probe with every
    // region shaded differently): whole table → row bands → column bands →
    // first/last column → first/last row → corner cells. So the header row
    // wins over the first column, and column bands over row bands (the
    // ECMA-376 text puts rows after columns there; Word unverified).

    struct TableLook {
        bool first_row = false, last_row = false, first_col = false, last_col = false;
        bool h_band = true, v_band = true;
    };

    // Without <w:tblLook>: header row, first column and row bands — what
    // Word gives a new table (val 04A0), and what LibreOffice assumes.
    static TableLook read_table_look(pugi::xml_node look) {
        if (!look) return TableLook{true, false, true, false, true, false};
        TableLook l;
        if (pugi::xml_attribute v = look.attribute("w:val")) {
            const unsigned long bits = std::strtoul(v.value(), nullptr, 16);
            l.first_row = (bits & 0x0020) != 0;
            l.last_row  = (bits & 0x0040) != 0;
            l.first_col = (bits & 0x0080) != 0;
            l.last_col  = (bits & 0x0100) != 0;
            l.h_band    = (bits & 0x0200) == 0;
            l.v_band    = (bits & 0x0400) == 0;
        }
        auto flag = [&](const char* name, bool& out, bool negated) {
            if (pugi::xml_attribute a = look.attribute(name)) {
                const std::string_view v = a.value();
                const bool on = v == "1" || v == "true" || v == "on";
                out = negated ? !on : on;
            }
        };
        flag("w:firstRow", l.first_row, false);
        flag("w:lastRow", l.last_row, false);
        flag("w:firstColumn", l.first_col, false);
        flag("w:lastColumn", l.last_col, false);
        flag("w:noHBand", l.h_band, true);
        flag("w:noVBand", l.v_band, true);
        return l;
    }

    struct CellPosition {
        std::size_t row = 0, rows = 1;    // in the table
        std::size_t cell = 0, cells = 1;  // in its row
        std::size_t grid_col = 0;
    };

    // The table style's formatting for the cell at `at`, and the border
    // visibility its conditional formatting gives it.
    CellStyle cell_style(pugi::xml_node style, const TableLook& look, const CellPosition& at,
                                StyleBorders& borders) {
        CellStyle cs;
        if (!style) return cs;
        // Which edges of a region a cell sits on decides whether a region's
        // top/left/bottom/right or its insideH/insideV border applies.
        enum class Region { Whole, Row, Column, Cell };
        auto add = [&](pugi::xml_node node, Region region) {
            if (!node) return;
            if (pugi::xml_node n = node.child("w:pPr")) cs.ppr.push_back(n);
            if (pugi::xml_node n = node.child("w:rPr")) cs.rpr.push_back(n);
            pugi::xml_node tc_pr = node.child("w:tcPr");
            if (!tc_pr) return;
            cs.tc_pr.push_back(tc_pr);
            pugi::xml_node tcb = tc_pr.child("w:tcBorders");
            if (!tcb) return;
            const bool first_row = at.row == 0, last_row = at.row + 1 == at.rows;
            const bool first_cell = at.cell == 0, last_cell = at.cell + 1 == at.cells;
            const bool top_edge    = region != Region::Column && region != Region::Whole ? true : first_row;
            const bool bottom_edge = region != Region::Column && region != Region::Whole ? true : last_row;
            const bool left_edge   = region != Region::Row && region != Region::Whole ? true : first_cell;
            const bool right_edge  = region != Region::Row && region != Region::Whole ? true : last_cell;
            auto side = [&](std::optional<Border>& out, bool edge, const char* outer, const char* outer_alt,
                            const char* inside) {
                pugi::xml_node b = edge ? tcb.child(outer) : tcb.child(inside);
                if (!b && edge && outer_alt != nullptr) b = tcb.child(outer_alt);
                if (b) out = parse_border(b, theme_);
            };
            side(borders.top, top_edge, "w:top", nullptr, "w:insideH");
            side(borders.bottom, bottom_edge, "w:bottom", nullptr, "w:insideH");
            side(borders.left, left_edge, "w:left", "w:start", "w:insideV");
            side(borders.right, right_edge, "w:right", "w:end", "w:insideV");
        };
        auto region = [&](const char* type) {
            for (pugi::xml_node n : style.children("w:tblStylePr")) {
                if (std::strcmp(n.attribute("w:type").value(), type) == 0) return n;
            }
            return pugi::xml_node{};
        };

        const bool header = look.first_row && at.row == 0;
        const bool total  = look.last_row && at.row + 1 == at.rows;
        const bool first  = look.first_col && at.cell == 0;
        const bool last   = look.last_col && at.cell + 1 == at.cells;
        pugi::xml_node style_pr = style.child("w:tblPr");
        const std::size_t row_band = std::max(1u, style_pr.child("w:tblStyleRowBandSize").attribute("w:val").as_uint(1));
        const std::size_t col_band = std::max(1u, style_pr.child("w:tblStyleColBandSize").attribute("w:val").as_uint(1));

        add(style, Region::Whole);  // the style's own pPr/rPr/tcPr
        add(region("wholeTable"), Region::Whole);
        if (look.h_band && !header) {
            const std::size_t r = at.row - (look.first_row ? 1 : 0);
            add(region((r / row_band) % 2 == 0 ? "band1Horz" : "band2Horz"), Region::Row);
        }
        if (look.v_band && !first) {
            const std::size_t c = at.grid_col - (look.first_col && at.grid_col > 0 ? 1 : 0);
            add(region((c / col_band) % 2 == 0 ? "band1Vert" : "band2Vert"), Region::Column);
        }
        if (first)  add(region("firstCol"), Region::Column);
        if (last)   add(region("lastCol"), Region::Column);
        if (header) add(region("firstRow"), Region::Row);
        if (total)  add(region("lastRow"), Region::Row);
        if (header && first) add(region("nwCell"), Region::Cell);
        if (header && last)  add(region("neCell"), Region::Cell);
        if (total && first)  add(region("swCell"), Region::Cell);
        if (total && last)   add(region("seCell"), Region::Cell);
        return cs;
    }

    TableLayout measure_table(pugi::xml_node tbl, double avail, bool force_fit,
                              const RelMap& rels, const PartMap& parts) {
        TableLayout t;
        const std::vector<pugi::xml_node> row_nodes = rows_of(tbl);
        if (row_nodes.empty()) return t;

        pugi::xml_node tbl_pr = tbl.child("w:tblPr");
        t.tbl_borders = tbl_pr.child("w:tblBorders");
        pugi::xml_node style = find_table_style(styles_doc_, tbl_pr.child("w:tblStyle").attribute("w:val").value());
        if (style) {
            t.style_borders  = style.child("w:tblPr").child("w:tblBorders");
            t.style_resolved = true;
        }
        const TableLook look = read_table_look(tbl_pr.child("w:tblLook"));
        CellMargins table_margins;
        apply_margins(style.child("w:tblPr").child("w:tblCellMar"), table_margins);
        apply_margins(tbl_pr.child("w:tblCellMar"), table_margins);

        const std::vector<double> col_widths = table_column_widths(tbl, avail, force_fit);
        for (double w : col_widths) t.width += w;
        const std::string_view jc = tbl_pr.child("w:jc").attribute("w:val").value();
        if (jc == "center")                   t.x_offset = std::max(0.0, (avail - t.width) / 2.0);
        else if (jc == "right" || jc == "end") t.x_offset = std::max(0.0, avail - t.width);
        else t.x_offset = twips_to_pt(tbl_pr.child("w:tblInd").attribute("w:w").as_llong(0));

        // Vertical merges: a <w:vMerge w:val="restart"> cell spans the
        // following rows whose cell at the same grid column continues it.
        struct MergeInfo {
            std::size_t grid_col = 0;
            bool        restart  = false;
            bool        cont     = false;
        };
        std::vector<std::vector<MergeInfo>> merge(row_nodes.size());
        for (std::size_t r = 0; r < row_nodes.size(); ++r) {
            std::size_t grid = 0;
            for (pugi::xml_node tc : cells_of(row_nodes[r])) {
                MergeInfo mi;
                mi.grid_col = grid;
                pugi::xml_node tc_pr = tc.child("w:tcPr");
                if (pugi::xml_node vm = tc_pr.child("w:vMerge")) {
                    mi.restart = std::strcmp(vm.attribute("w:val").value(), "restart") == 0;
                    mi.cont    = !mi.restart;
                }
                grid += std::max(1u, tc_pr.child("w:gridSpan").attribute("w:val").as_uint(1));
                merge[r].push_back(mi);
            }
        }
        auto continues_at = [&](std::size_t r, std::size_t grid_col) {
            for (const MergeInfo& mi : merge[r]) {
                if (mi.grid_col == grid_col) return mi.cont;
            }
            return false;
        };

        for (std::size_t r = 0; r < row_nodes.size(); ++r) {
            pugi::xml_node tr = row_nodes[r];
            RowLayout row;
            pugi::xml_node tr_pr = tr.child("w:trPr");
            row.header = on_off(tr_pr.child("w:tblHeader"));
            const std::vector<pugi::xml_node> cells = cells_of(tr);
            const std::vector<double> widths = cell_widths_for_row(cells, col_widths);
            for (std::size_t i = 0; i < cells.size(); ++i) {
                CellLayout c;
                c.tc    = cells[i];
                c.width = widths[i];
                c.grid_col = merge[r][i].grid_col;
                c.covered  = merge[r][i].cont;
                if (merge[r][i].restart) {
                    while (r + c.rowspan < row_nodes.size() && continues_at(r + c.rowspan, merge[r][i].grid_col)) {
                        ++c.rowspan;
                    }
                }
                pugi::xml_node tc_pr = cells[i].child("w:tcPr");
                c.margins = table_margins;
                apply_margins(tc_pr.child("w:tcMar"), c.margins);
                const CellStyle cs = cell_style(style, look, {r, row_nodes.size(), i, cells.size(), c.grid_col},
                                                c.style_borders);
                for (pugi::xml_node style_tc_pr : cs.tc_pr) {
                    if (pugi::xml_node shd = style_tc_pr.child("w:shd")) {
                        if (auto fill = shading_fill(shd, theme_)) c.fill = fill;
                        else if (std::strcmp(shd.attribute("w:fill").value(), "auto") == 0 ||
                                 std::strcmp(shd.attribute("w:val").value(), "nil") == 0) c.fill.reset();
                    }
                }
                if (auto fill = shading_fill(tc_pr.child("w:shd"), theme_)) c.fill = fill;
                const std::string_view va = tc_pr.child("w:vAlign").attribute("w:val").value();
                c.valign = va == "center" ? CellLayout::VAlign::Center
                         : va == "bottom" ? CellLayout::VAlign::Bottom : CellLayout::VAlign::Top;
                if (!c.covered) {
                    c.content = classify_cell(cells[i], std::max(c.width - c.margins.left - c.margins.right, 1.0),
                                              style ? &cs : nullptr, rels, parts);
                }
                // A merged cell's height is spread over its rows below.
                if (c.rowspan == 1 && !c.covered) {
                    row.height = std::max(row.height, c.content.height + c.margins.top + c.margins.bottom);
                }
                row.cells.push_back(std::move(c));
            }
            // A row always holds at least one line of text.
            row.height = std::max(row.height, metrics_of(styles_.mark_style(pugi::xml_node{})).line);
            if (pugi::xml_node h = tr_pr.child("w:trHeight")) {
                const double v = twips_to_pt(h.attribute("w:val").as_llong(0));
                if (std::strcmp(h.attribute("w:hRule").value(), "exact") == 0) row.height = v;
                else row.height = std::max(row.height, v);
            }
            t.rows.push_back(std::move(row));
        }

        // Grow the last row of a merged area if the merged cell needs more
        // room than its rows give it. An area taller than a page is drawn
        // unmerged instead (the table breaks between its rows).
        const double page_body = pm_.height_pt - pm_.margin_top_pt - pm_.margin_bottom_pt;
        for (std::size_t r = 0; r < t.rows.size(); ++r) {
            for (CellLayout& c : t.rows[r].cells) {
                if (c.rowspan <= 1) continue;
                double have = 0.0;
                for (std::size_t k = r; k < r + c.rowspan; ++k) have += t.rows[k].height;
                const double need = c.content.height + c.margins.top + c.margins.bottom;
                if (std::max(have, need) > page_body * 0.9) {
                    for (std::size_t k = r + 1; k < r + c.rowspan; ++k) {
                        for (CellLayout& other : t.rows[k].cells) {
                            if (other.covered && other.grid_col == c.grid_col) other.covered = false;
                        }
                    }
                    c.rowspan = 1;
                    t.rows[r].height = std::max(t.rows[r].height, need);
                    continue;
                }
                if (need > have) t.rows[r + c.rowspan - 1].height += need - have;
            }
        }
        return t;
    }

    // Rows [ri, result) must stay on one page: merged cells starting in
    // them reach down to result - 1.
    static std::size_t merged_block_end(const TableLayout& t, std::size_t ri) {
        std::size_t end = ri + 1;
        for (std::size_t k = ri; k < end && k < t.rows.size(); ++k) {
            for (const CellLayout& c : t.rows[k].cells) end = std::max(end, k + c.rowspan);
        }
        return std::min(end, t.rows.size());
    }

    void draw_row(const TableLayout& t, std::size_t ri, double x, double top) {
        if (dry_run_) return;
        const RowLayout& row = t.rows[ri];
        for (std::size_t i = 0; i < row.cells.size(); ++i) {
            const CellLayout& c = row.cells[i];
            if (c.covered) { x += c.width; continue; }  // drawn with its merged cell above
            record_bookmarks(c.tc);
            double height = 0.0;
            for (std::size_t k = ri; k < ri + c.rowspan && k < t.rows.size(); ++k) height += t.rows[k].height;
            if (c.fill) {
                painter_.GraphicsState.SetNonStrokingColor(pdf_color(*c.fill));
                painter_.DrawRectangle(x, top - height, c.width, height, PoDoFo::PdfPathDrawMode::Fill);
                painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            }
            const double inner_h = height - c.margins.top - c.margins.bottom;
            double offset = 0.0;
            if (c.valign == CellLayout::VAlign::Center)      offset = std::max(0.0, (inner_h - c.content.height) / 2.0);
            else if (c.valign == CellLayout::VAlign::Bottom) offset = std::max(0.0, inner_h - c.content.height);
            draw_cell_content(c.content, x + c.margins.left, top - c.margins.top - offset,
                              std::max(c.width - c.margins.left - c.margins.right, 1.0));
            CellBorderSides sides = borders_of(t, ri, i);
            // A shared edge is drawn once more by the cell below / to the
            // right, with whichever of the two borders wins — over the
            // loser the neighbour already drew.
            if (auto above = cell_above(t, ri, c.grid_col)) {
                sides.top = stronger(sides.top, borders_of(t, above->first, above->second).bottom);
            }
            if (i > 0) sides.left = stronger(sides.left, borders_of(t, ri, i - 1).right);
            draw_cell_border(x, top - height, c.width, height, sides);
            x += c.width;
        }
    }

    CellBorderSides borders_of(const TableLayout& t, std::size_t ri, std::size_t i) const {
        const RowLayout& row = t.rows[ri];
        const CellLayout& c = row.cells[i];
        return resolve_cell_borders(theme_, c.tc, t.tbl_borders, t.style_borders, t.style_resolved, c.style_borders,
                                    ri == 0, ri + c.rowspan >= t.rows.size(), i == 0, i == row.cells.size() - 1);
    }

    // The drawn cell (row, index) directly above grid column `grid_col` of
    // row `ri` — the start of a vertical merge rather than its covered part.
    static std::optional<std::pair<std::size_t, std::size_t>> cell_above(const TableLayout& t, std::size_t ri,
                                                                         std::size_t grid_col) {
        for (std::size_t r = ri; r-- > 0;) {
            const std::vector<CellLayout>& cells = t.rows[r].cells;
            for (std::size_t i = 0; i < cells.size(); ++i) {
                const std::size_t span =
                    std::max(1u, cells[i].tc.child("w:tcPr").child("w:gridSpan").attribute("w:val").as_uint(1));
                if (grid_col >= cells[i].grid_col && grid_col < cells[i].grid_col + span) {
                    if (!cells[i].covered) return std::make_pair(r, i);
                    break;  // covered: keep looking upwards for its merge start
                }
            }
        }
        return std::nullopt;
    }

    PageFields current_fields() const {
        if (fields_.page > 0) return fields_;
        PageFields f{pages_.empty() ? 0 : pages_.back().number, 0, 0,
                     pages_.empty() ? std::string() : sections_[pages_.back().section].page_number_format, nullptr};
        if (known_ != nullptr) {
            f.total = known_->total;
            if (const auto it = known_->section_pages.find(section_); it != known_->section_pages.end()) {
                f.section_total = it->second;
            }
            f.bookmark_pages = &known_->bookmark_pages;
        }
        return f;
    }

    // Notes the page of every bookmark starting in `node` (a paragraph or
    // a table cell being drawn on the current page); the first place wins.
    void record_bookmarks(pugi::xml_node node) {
        if (dry_run_ || no_page_break_ || pages_.empty()) return;
        for (pugi::xml_node child : node.children()) {
            if (std::strcmp(child.name(), "w:bookmarkStart") == 0) {
                bookmark_pages_.emplace(child.attribute("w:name").value(),
                                        PageNumber{pages_.back().number,
                                                   sections_[pages_.back().section].page_number_format});
            } else if (child.first_child()) {
                record_bookmarks(child);
            }
        }
    }

    // ── Hidden text (<w:vanish>) ────────────────────────────────────────────

    std::vector<Fragment> fragments_of(pugi::xml_node p) const {
        return ParagraphText(current_fields(), [&](pugi::xml_node r) { return styles_.run_style(p, r, cell_style_); },
                             styles_.mark_style(p, cell_style_))
            .run(p);
    }

    // Pictures and charts of `p` laid out in the text flow: the inline ones
    // (floating ones are placed by place_floats()), all of them in a cell.
    std::vector<pugi::xml_node> flow_drawings(pugi::xml_node p) const {
        std::vector<pugi::xml_node> out;
        for (pugi::xml_node d : visible_drawings(p)) {
            if (in_cell_ || !d.child("wp:anchor")) out.push_back(d);
        }
        return out;
    }

    // Pictures and charts of `p` that are printed: not inside a hidden run.
    std::vector<pugi::xml_node> visible_drawings(pugi::xml_node p) const {
        std::vector<pugi::xml_node> out;
        for (pugi::xml_node drawing : drawings_of(p)) {
            pugi::xml_node r = drawing.parent();
            while (r && r != p && std::strcmp(r.name(), "w:r") != 0) r = r.parent();
            if (!r || r == p || !styles_.run_style(p, r, cell_style_).hidden) out.push_back(drawing);
        }
        return out;
    }

    // A paragraph whose mark and whole content are hidden takes no room at
    // all — no line, no spacing, no list number — as in Word. (With a
    // hidden mark but visible content Word joins it to the next paragraph;
    // that is drawn as a paragraph of its own here.)
    bool hidden_paragraph(pugi::xml_node p) const {
        if (std::strcmp(p.name(), "w:p") != 0 || !styles_.mark_style(p, cell_style_).hidden) return false;
        return fragments_of(p).empty() && visible_drawings(p).empty();
    }

    // Height `blocks` take when laid out at full content width.
    double measure_blocks(const std::vector<pugi::xml_node>& blocks, const RelMap& rels,
                          const PartMap& parts) {
        const double    saved_cursor = cursor_y_;
        const bool      saved_dry    = dry_run_;
        const FlowState saved_flow   = flow_;
        dry_run_  = true;
        cursor_y_ = 0.0;
        flow_.reset();
        draw_blocks(blocks, rels, parts);
        const double height = -cursor_y_;
        cursor_y_ = saved_cursor;
        dry_run_  = saved_dry;
        flow_     = saved_flow;
        return height;
    }

    // Lays out one <w:tc>'s content at `width` points wide: paragraphs,
    // pictures, charts and nested tables stacked in document order.
    // `style`: what the table style gives this cell (nullptr: no style).
    CellContent classify_cell(pugi::xml_node tc, double width, const CellStyle* style,
                              const RelMap& rels, const PartMap& parts) {
        const CellStyle* const outer_style = cell_style_;  // a nested table's cell sits in another's
        cell_style_ = style;
        struct Restore {
            const CellStyle*& slot;
            const CellStyle*  value;
            ~Restore() { slot = value; }
        } restore{cell_style_, outer_style};
        const bool outer_in_cell = in_cell_;
        in_cell_ = true;
        struct RestoreFlag {
            bool& slot;
            bool  value;
            ~RestoreFlag() { slot = value; }
        } restore_flag{in_cell_, outer_in_cell};
        CellContent c;
        FlowState flow;
        double y = 0.0;
        auto sized = [&](CellPart& part, double w, double h) {
            part.width  = w;
            part.height = h;
            if (w > width && w > 0.0) {
                part.height = h * width / w;
                part.width  = width;
            }
        };

        std::vector<pugi::xml_node> blocks;
        for (pugi::xml_node block : blocks_of(tc)) {
            if (!hidden_paragraph(block)) blocks.push_back(block);
        }
        for (std::size_t bi = 0; bi < blocks.size(); ++bi) {
            const pugi::xml_node block = blocks[bi];
            if (std::strcmp(block.name(), "w:tbl") == 0) {
                y += flow.pending_after;
                flow.reset();
                CellPart part;
                part.kind   = CellPart::Kind::NestedTable;
                part.nested = std::make_shared<TableLayout>(measure_table(block, width, /*force_fit*/true, rels, parts));
                part.width  = width;
                part.height = part.nested->total_height();
                part.y      = y;
                y += part.height;
                c.parts.push_back(std::move(part));
                continue;
            }

            LaidParagraph lp = layout_paragraph(block, width);
            const std::optional<ParaProps> prev_pp = props_of(bi > 0 ? blocks[bi - 1] : pugi::xml_node{});
            const std::optional<ParaProps> next_pp = props_of(bi + 1 < blocks.size() ? blocks[bi + 1] : pugi::xml_node{});
            const ParaBox box(lp.props, prev_pp ? &*prev_pp : nullptr, next_pp ? &*next_pp : nullptr);
            const double gap = flow.gap_before(lp.props);
            y += gap;
            const double box_top = y;
            y += box.top;
            for (pugi::xml_node drawing : visible_drawings(block)) {
                CellPart part;
                part.align = lp.props.align;
                if (pugi::xml_node chart_ref = find_descendant(drawing, "c:chart")) {
                    part.kind  = CellPart::Kind::Chart;
                    part.chart = resolve_chart(drawing, chart_ref, rels, parts, theme_);
                    sized(part, part.chart.width_pt, part.chart.height_pt);
                } else {
                    part.kind  = CellPart::Kind::Image;
                    part.image = resolve_image(drawing, rels, parts);
                    sized(part, part.image.width_pt, part.image.height_pt);
                }
                part.y = y;
                y += part.height;
                c.parts.push_back(std::move(part));
            }
            if (!lp.lines.empty()) {
                CellPart part;
                part.kind   = CellPart::Kind::Paragraph;
                part.height = lp.lines_height();
                part.y      = y;
                part.box       = box;
                part.box_above = y - box_top;
                part.box_below = box.bottom;
                part.box_gap   = box.first ? 0.0 : gap;
                part.paragraph = std::move(lp);
                y += part.height + box.bottom;
                flow.finished(part.paragraph.props);
                c.parts.push_back(std::move(part));
            } else {
                y += box.bottom;
                flow.finished(lp.props);
            }
        }
        c.height = y + flow.pending_after;
        return c;
    }

    // Draws the visible sides as independent line segments rather than one
    // rectangle — two adjacent cells that disagree about their shared edge
    // (a header cell's own <w:tcBorders> claiming a bottom line the row
    // below doesn't repeat as its own top, say) both still get an honest
    // rendering; draw_row() redraws a shared edge with the winning border.
    void draw_cell_border(double x, double y, double w, double h, const CellBorderSides& sides) {
        stroke_border(sides.top,    x, y + h, x + w, y + h);
        stroke_border(sides.bottom, x, y, x + w, y);
        stroke_border(sides.left,   x, y, x, y + h);
        stroke_border(sides.right,  x + w, y, x + w, y + h);
    }

    // One border along a horizontal or vertical edge.
    void stroke_border(const Border& b, double x1, double y1, double x2, double y2) {
        if (!b.visible()) return;
        painter_.GraphicsState.SetStrokingColor(pdf_color(b.color));
        painter_.GraphicsState.SetLineWidth(b.width);
        const double w = b.width;
        std::vector<double> dash;
        switch (b.style) {
            case Border::Style::Dotted:  dash = {std::max(w, 0.5), std::max(2.0 * w, 1.5)}; break;
            case Border::Style::Dashed:  dash = {std::max(3.0 * w, 3.0), std::max(2.0 * w, 2.0)}; break;
            case Border::Style::DotDash: dash = {std::max(3.0 * w, 3.0), std::max(2.0 * w, 1.5),
                                                 std::max(w, 0.5), std::max(2.0 * w, 1.5)}; break;
            default: break;
        }
        if (!dash.empty()) painter_.SetStrokeStyle(PoDoFo::cspan<double>(dash.data(), dash.size()), 0.0);
        if (b.style == Border::Style::Double) {
            // Two lines of the border's width, a gap of the same (at least
            // 0.75 pt) between them, centred on the edge.
            const double off = (w + std::max(w, 0.75)) / 2.0;
            const bool horizontal = y1 == y2;
            const double dx = horizontal ? 0.0 : off, dy = horizontal ? off : 0.0;
            painter_.DrawLine(x1 - dx, y1 - dy, x2 - dx, y2 - dy);
            painter_.DrawLine(x1 + dx, y1 + dy, x2 + dx, y2 + dy);
        } else {
            painter_.DrawLine(x1, y1, x2, y2);
        }
        if (!dash.empty()) painter_.SetStrokeStyle(PoDoFo::PdfStrokeStyle::Solid);
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetLineWidth(1.0);
    }

    // Draws one cell's parts at the given content-box origin (inside the
    // cell margins).
    void draw_cell_content(const CellContent& c, double x, double top_y, double width) {
        for (const CellPart& part : c.parts) {
            const double y = top_y - part.y;
            double dx = 0.0;
            if (part.align == ParaProps::Align::Center)     dx = std::max(0.0, (width - part.width) / 2.0);
            else if (part.align == ParaProps::Align::Right) dx = std::max(0.0, width - part.width);
            switch (part.kind) {
                case CellPart::Kind::Paragraph: {
                    const ParaProps& pp = part.paragraph.props;
                    if (pp.boxed()) {
                        const double top_line = y + part.box_above;
                        const double bottom   = y - part.height - part.box_below;
                        fill_box(pp, x, width, top_line + part.box_gap, bottom);
                        stroke_box(pp, part.box, x, width, top_line + part.box_gap, bottom, top_line, true);
                    }
                    double ty = y;
                    for (const Line& line : part.paragraph.lines) {
                        draw_text_line(line, part.paragraph.props, x, ty);
                        ty -= line.height;
                    }
                    break;
                }
                case CellPart::Kind::Image: {
                    std::unique_ptr<PoDoFo::PdfImage> image = doc_.CreateImage();
                    image->LoadFromBuffer(
                        PoDoFo::bufferview(part.image.png_bytes.data(), part.image.png_bytes.size()));
                    const double scale_x = part.width  / static_cast<double>(image->GetWidth());
                    const double scale_y = part.height / static_cast<double>(image->GetHeight());
                    painter_.DrawImage(*image, x + dx, y - part.height, scale_x, scale_y);
                    break;
                }
                case CellPart::Kind::Chart:
                    // No legend — cells are usually too narrow to spare the
                    // extra row, and the category/series names are already
                    // visible via the rest of the cell's own table row/column
                    // headers in every real template this was tested against.
                    draw_chart_in_box(part.chart, x + dx, y - part.height, part.width, part.height);
                    break;
                case CellPart::Kind::NestedTable: {
                    double ry = y;
                    for (std::size_t ri = 0; ri < part.nested->rows.size(); ++ri) {
                        draw_row(*part.nested, ri, x + part.nested->x_offset, ry);
                        ry -= part.nested->rows[ri].height;
                    }
                    break;
                }
            }
        }
    }

    double content_width() const {
        return pm_.width_pt - pm_.margin_left_pt - pm_.margin_right_pt;
    }

    void ensure_space(double needed) {
        if (dry_run_ || no_page_break_) return;
        if (cursor_y_ - needed < bottom_limit_ && cursor_y_ < page_top_) new_page();
        if (avoid_bands(needed) && cursor_y_ - needed < bottom_limit_) new_page();
    }

    // Moves the cursor below any floating object's band that `needed`
    // points from here would run into; whether it moved.
    bool avoid_bands(double needed) {
        bool moved = false;
        for (bool again = true; again;) {
            again = false;
            for (const Band& b : bands_) {
                if (cursor_y_ > b.bottom + 0.01 && cursor_y_ - needed < b.top) {
                    cursor_y_ = b.bottom;
                    moved = again = true;
                }
            }
        }
        return moved;
    }

    // Vertical space between blocks; dropped at a page break.
    void advance(double gap) {
        if (gap <= 0.0) return;
        if (!dry_run_ && !no_page_break_ && cursor_y_ - gap < bottom_limit_) {
            new_page();
            return;
        }
        cursor_y_ -= gap;
    }

    // Manual page break: start a new page unless nothing was drawn on the
    // current one yet.
    void break_page() {
        if (dry_run_ || no_page_break_) return;
        if (cursor_y_ < page_top_) new_page();
    }

    // Plain single-style text (chart labels, legends).
    void draw_line(PoDoFo::PdfFont& font, double size, const std::string& text,
                   double x, double y) {
        if (dry_run_ || text.empty()) return;
        // DrawText() manages its own BT/ET text object internally in this
        // PoDoFo version (BeginText/EndText are private — meant for a
        // lower-level multi-call text object API this renderer doesn't need).
        const std::string line = sanitize_single_line(text);
        painter_.TextState.SetFont(fonts_.for_text(font, line), size);
        painter_.DrawText(line, x, y);
    }

    // Width of chart/legend text as draw_line() will draw it.
    double label_width(PoDoFo::PdfFont& font, double size, const std::string& text) {
        const std::string line = sanitize_single_line(text);
        return text_width(fonts_.for_text(font, line), size, line);
    }

    // Horizontal position of a `w`-wide block in a paragraph with `pp`.
    double block_x(double w, const ParaProps& pp) const {
        const double left  = pm_.margin_left_pt + pp.ind_left;
        const double avail = content_width() - pp.ind_left - pp.ind_right;
        if (pp.align == ParaProps::Align::Center) return left + std::max(0.0, (avail - w) / 2.0);
        if (pp.align == ParaProps::Align::Right)  return left + std::max(0.0, avail - w);
        return left;
    }

    void draw_image(const ImageBlock& img, const ParaProps& pp) {
        // Scaled down (aspect kept) to the content width, like Word.
        double w = img.width_pt;
        double h = img.height_pt;
        if (w > content_width() && w > 0.0) {
            h *= content_width() / w;
            w  = content_width();
        }
        ensure_space(h);
        if (!dry_run_) {
            std::unique_ptr<PoDoFo::PdfImage> image = doc_.CreateImage();
            image->LoadFromBuffer(PoDoFo::bufferview(img.png_bytes.data(), img.png_bytes.size()));
            const double scale_x = w / static_cast<double>(image->GetWidth());
            const double scale_y = h / static_cast<double>(image->GetHeight());
            painter_.DrawImage(*image, block_x(w, pp), cursor_y_ - h, scale_x, scale_y);
        }
        cursor_y_ -= h;
    }

    // Resolves and renders a chart at page scale (bar/line/area/pie/
    // doughnut, 3D variants rendered flat) or throws NotImplemented —
    // see resolve_chart(). This is the page-flow entry point (with a
    // legend row); a chart inside a table cell goes through
    // draw_cell_content() instead, which skips the legend.
    void draw_chart_block(pugi::xml_node drawing, pugi::xml_node chart_ref,
                          const RelMap& rels, const PartMap& parts, const ParaProps& pp) {
        const ResolvedChart rc = resolve_chart(drawing, chart_ref, rels, parts, theme_);
        const double width_pt  = std::min(rc.width_pt, content_width());
        const double height_pt = rc.height_pt;

        constexpr double kLegendRowHeight = 16.0;
        ensure_space(height_pt + kLegendRowHeight);

        const double chart_top = cursor_y_;
        if (!dry_run_) {
            const double x = block_x(width_pt, pp);
            draw_chart_in_box(rc, x, chart_top - height_pt, width_pt, height_pt);
            if (rc.kind == ChartKind::Pie || rc.kind == ChartKind::Doughnut) {
                draw_category_legend(rc.data, x, chart_top - height_pt - 4.0);
            } else {
                draw_chart_legend(rc.data, x, chart_top - height_pt - 4.0);
            }
        }
        cursor_y_ = chart_top - height_pt - kLegendRowHeight;
    }

    // Title (if any) on top, then the plot in the remaining box.
    void draw_chart_in_box(const ResolvedChart& rc, double x, double y, double w, double h) {
        if (dry_run_) return;
        constexpr double kTitleSize = 10.0;
        if (!rc.title.empty()) {
            PoDoFo::PdfTextState title_state;
            title_state.Font     = &bold_;
            title_state.FontSize = kTitleSize;
            const double title_w = label_width(bold_, title_state.FontSize, rc.title);
            draw_line(bold_, kTitleSize, rc.title,
                      x + std::max((w - title_w) / 2.0, 0.0), y + h - kTitleSize);
            h -= kTitleSize * 1.6;
        }
        if (rc.kind == ChartKind::Pie || rc.kind == ChartKind::Doughnut) {
            draw_pie_chart_plot(rc.kind == ChartKind::Doughnut, rc.hole_frac, rc.data, x, y, w, h);
        } else {
            draw_chart_plot(rc, x, y, w, h);
        }
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

        std::vector<PendingLabel> labels;
        const ChartSeriesData& ser = data.series.front();

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
                append_arc(path, cx, cy, radius, a_start, a_end);
                path.AddLineTo(cx + hole_r * std::cos(a_end), cy + hole_r * std::sin(a_end));
                append_arc(path, cx, cy, hole_r, a_end, a_start);
                path.Close();
            } else {
                path.MoveTo(cx, cy);
                path.AddLineTo(cx + radius * std::cos(a_start), cy + radius * std::sin(a_start));
                append_arc(path, cx, cy, radius, a_start, a_end);
                path.Close();
            }
            painter_.GraphicsState.SetNonStrokingColor(pdf_color(point_color(data, i)));
            painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Fill);

            // Data labels (<c:dLbls>) — none unless the chart asks for
            // them. Inside at the middle (ctr), near the rim (inEnd),
            // outside it (outEnd); best fit and the default: inside.
            const DataLabels& dl = labels_at(ser, i);
            if (!dl.shown()) continue;
            const std::string& pos = dl.position;
            LabelPlace place = LabelPlace::Center;
            double label_r = hole_r > 0.0 ? (hole_r + radius) / 2.0 : radius * 0.65;
            if (hole_r <= 0.0 && pos == "ctr") label_r = radius * 0.5;
            if (hole_r <= 0.0 && pos == "inEnd") label_r = radius * 0.8;
            if (pos == "outEnd") {
                label_r = radius + 4.0;
                place = std::cos(a_mid) >= 0.0 ? LabelPlace::Right : LabelPlace::Left;
            }
            labels.push_back({data_label_text(dl, ser, data.categories[i], values[i], frac),
                              cx + label_r * std::cos(a_mid), cy + label_r * std::sin(a_mid), place, &dl});
        }

        draw_data_labels(labels);
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
            const double        text_w = label_width(regular_, state.FontSize, name);
            const double        item_w = kSwatch + 4.0 + text_w + kGap * 2.0;
            if (cursor_x + item_w > max_x && cursor_x > x) break;

            painter_.GraphicsState.SetNonStrokingColor(pdf_color(point_color(data, i)));
            painter_.DrawRectangle(cursor_x, y, kSwatch, kSwatch, PoDoFo::PdfPathDrawMode::Fill);
            painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            draw_line(regular_, kFontSize, name, cursor_x + kSwatch + 4.0, y + 1.0);
            cursor_x += item_w;
        }
    }

    // Value of series `si` at category `i` as drawn: normalized to its
    // category total in a 100 % stacked group.
    static double group_value(const ChartData& data, const ChartGroup& g, std::size_t si, std::size_t i) {
        const auto& values = data.series[si].values;
        const double v = i < values.size() ? values[i] : 0.0;
        if (g.grouping != Grouping::PercentStacked) return v;
        double total = 0.0;
        for (std::size_t s2 : g.series) {
            const auto& vs = data.series[s2].values;
            total += std::abs(i < vs.size() ? vs[i] : 0.0);
        }
        return total > 0.0 ? v / total : 0.0;
    }

    // Value range the groups on one axis need (stacks summed per category).
    static std::pair<double, double> axis_range(const ChartData& data, bool secondary) {
        double lo = 0.0, hi = 0.0;
        const std::size_t ncat = data.categories.size();
        for (const ChartGroup& g : data.groups) {
            if (g.secondary != secondary) continue;
            for (std::size_t i = 0; i < ncat; ++i) {
                double pos = 0.0, neg = 0.0;
                for (std::size_t si : g.series) {
                    const double v = group_value(data, g, si, i);
                    if (g.grouping == Grouping::Clustered) {
                        lo = std::min(lo, v);
                        hi = std::max(hi, v);
                    } else if (v >= 0.0) {
                        pos += v;
                    } else {
                        neg += v;
                    }
                }
                lo = std::min(lo, neg);
                hi = std::max(hi, pos);
            }
        }
        return {lo, hi};
    }

    struct AxisTicks {
        AxisScale                                   scale;
        std::vector<std::pair<double, std::string>> ticks;
        double                                      widest = 0.0;
        bool                                        percent = false;
    };

    AxisTicks axis_ticks(const ChartData& data, const ChartAxis& axis, bool secondary, double label_size) {
        AxisTicks at;
        at.percent = true;
        bool any = false;
        for (const ChartGroup& g : data.groups) {
            if (g.secondary != secondary) continue;
            any = true;
            at.percent = at.percent && g.grouping == Grouping::PercentStacked;
        }
        at.percent = at.percent && any;
        const auto [lo, hi] = axis_range(data, secondary);
        at.scale = nice_axis_scale(lo, hi);
        if (axis.min || axis.max) {
            // Explicit bounds: keep them, pick a nice step for the span.
            const double mn = axis.min.value_or(at.scale.min);
            const double mx = axis.max.value_or(at.scale.max);
            if (mx > mn) {
                AxisScale sc = nice_axis_scale(0.0, mx - mn);
                sc.min = mn;
                sc.max = mx;
                at.scale = sc;
            }
        }
        const int count = static_cast<int>(std::floor((at.scale.max - at.scale.min) / at.scale.step + 1e-6));
        for (int i = 0; i <= count; ++i) {
            const double v = at.scale.min + at.scale.step * i;
            const std::string label = at.percent
                ? format_axis_value(v * 100.0, std::max(0, at.scale.decimals - 2)) + "%"
                : format_axis_value(v, at.scale.decimals);
            at.ticks.emplace_back(v, label);
            at.widest = std::max(at.widest, label_width(regular_, label_size, label));
        }
        return at;
    }

    // Draws an axis chart (bar/line/area groups, any mix; stacked and 100 %
    // stacked; vertical or horizontal bars; secondary value axis) in the box
    // (x, y)-(x+w, y+h): gridlines, axis labels and titles, then areas,
    // bars and lines in that order, as Excel layers them.
    void draw_chart_plot(const ResolvedChart& rc, double x, double y, double w, double h) {
        constexpr double kLabelSize = 7.0;
        constexpr double kTitleSize = 8.0;
        const ChartData& data = rc.data;
        const std::size_t ncat = data.categories.size();
        const bool horiz = rc.horizontal;

        const AxisTicks primary   = axis_ticks(data, rc.val_axis, false, kLabelSize);
        const AxisTicks secondary = rc.has_secondary ? axis_ticks(data, rc.val2_axis, true, kLabelSize) : AxisTicks{};

        double widest_cat = 0.0;
        for (const std::string& c : data.categories) widest_cat = std::max(widest_cat, label_width(regular_, kLabelSize, c));

        // Paddings around the plot rectangle.
        double left, right, bottom, top;
        if (horiz) {
            left   = std::min(widest_cat + 8.0, w * 0.4);
            right  = 8.0;
            bottom = 14.0 + (rc.val_axis.title.empty() ? 0.0 : kTitleSize + 4.0);
            top    = 4.0 + (rc.cat_axis.title.empty() ? 0.0 : kTitleSize + 4.0);
        } else {
            left   = std::max(24.0, primary.widest + 8.0);
            right  = rc.has_secondary ? secondary.widest + 8.0 : 4.0;
            bottom = 16.0 + (rc.cat_axis.title.empty() ? 0.0 : kTitleSize + 4.0);
            top    = 4.0 + ((rc.val_axis.title.empty() && rc.val2_axis.title.empty()) ? 0.0 : kTitleSize + 4.0);
        }
        const double plot_x = x + left;
        const double plot_y = y + bottom;
        const double plot_w = std::max(w - left - right, 1.0);
        const double plot_h = std::max(h - bottom - top, 1.0);

        // Axis titles, horizontal.
        if (horiz) {
            if (!rc.cat_axis.title.empty()) draw_line(regular_, kTitleSize, rc.cat_axis.title, x, y + h - kTitleSize);
            if (!rc.val_axis.title.empty()) {
                const double tw = label_width(regular_, kTitleSize, rc.val_axis.title);
                draw_line(regular_, kTitleSize, rc.val_axis.title, plot_x + std::max((plot_w - tw) / 2.0, 0.0), y + 2.0);
            }
        } else {
            if (!rc.val_axis.title.empty()) draw_line(regular_, kTitleSize, rc.val_axis.title, x, y + h - kTitleSize);
            if (!rc.val2_axis.title.empty()) {
                const double tw = label_width(regular_, kTitleSize, rc.val2_axis.title);
                draw_line(regular_, kTitleSize, rc.val2_axis.title, x + w - tw, y + h - kTitleSize);
            }
            if (!rc.cat_axis.title.empty()) {
                const double tw = label_width(regular_, kTitleSize, rc.cat_axis.title);
                draw_line(regular_, kTitleSize, rc.cat_axis.title, plot_x + std::max((plot_w - tw) / 2.0, 0.0), y + 2.0);
            }
        }

        // Chart space: c = along the category axis, v = along the value axis.
        const double cat_len = horiz ? plot_h : plot_w;
        const double val_len = horiz ? plot_w : plot_h;
        const double slot    = cat_len / static_cast<double>(std::max<std::size_t>(ncat, 1));
        auto cat_center = [&](std::size_t i) {
            const std::size_t k = rc.cat_axis.reversed ? ncat - 1 - i : i;
            return (static_cast<double>(k) + 0.5) * slot;
        };
        auto vpos = [&](const AxisScale& sc, double v) {
            return (std::clamp(v, sc.min, sc.max) - sc.min) / (sc.max - sc.min) * val_len;
        };
        auto page_x = [&](double c, double v) { return horiz ? plot_x + v : plot_x + c; };
        auto page_y = [&](double c, double v) { return horiz ? plot_y + c : plot_y + v; };
        auto fill_rect = [&](double c0, double c1, double v0, double v1) {
            const double x0 = page_x(std::min(c0, c1), std::min(v0, v1));
            const double y0 = page_y(std::min(c0, c1), std::min(v0, v1));
            const double dc = std::abs(c1 - c0);
            const double dv = std::max(std::abs(v1 - v0), 0.5);
            painter_.DrawRectangle(x0, y0, horiz ? dv : dc, horiz ? dc : dv, PoDoFo::PdfPathDrawMode::Fill);
        };

        // Gridlines and value labels (primary), secondary labels on the right.
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0.75, 0.75, 0.75));
        painter_.GraphicsState.SetLineWidth(0.5);
        for (const auto& [v, label] : primary.ticks) {
            const double p = vpos(primary.scale, v);
            painter_.DrawLine(page_x(0, p), page_y(0, p), page_x(cat_len, p), page_y(cat_len, p));
            if (rc.val_axis.deleted) continue;
            const double lw = label_width(regular_, kLabelSize, label);
            if (horiz) draw_line(regular_, kLabelSize, label, plot_x + p - lw / 2.0, plot_y - 10.0);
            else       draw_line(regular_, kLabelSize, label, plot_x - lw - 4.0, plot_y + p - kLabelSize * 0.3);
        }
        if (rc.has_secondary && !rc.val2_axis.deleted && !horiz) {
            for (const auto& [v, label] : secondary.ticks) {
                const double p = vpos(secondary.scale, v);
                draw_line(regular_, kLabelSize, label, plot_x + plot_w + 4.0, plot_y + p - kLabelSize * 0.3);
            }
        }

        // Category labels; thinned out when they'd overlap.
        if (!rc.cat_axis.deleted && ncat > 0) {
            const double room = horiz ? kLabelSize * 1.3 : widest_cat + 4.0;
            const std::size_t every = std::max<std::size_t>(1, static_cast<std::size_t>(std::ceil(room / std::max(slot, 1.0))));
            for (std::size_t i = 0; i < ncat; i += every) {
                const double c  = cat_center(i);
                const double lw = label_width(regular_, kLabelSize, data.categories[i]);
                if (horiz) draw_line(regular_, kLabelSize, data.categories[i], plot_x - lw - 4.0, plot_y + c - kLabelSize * 0.3);
                else       draw_line(regular_, kLabelSize, data.categories[i], plot_x + c - lw / 2.0, plot_y - 10.0);
            }
        }

        // Axis lines: value axis along the plot edge, category axis at 0.
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetLineWidth(0.75);
        const double zero = vpos(primary.scale, 0.0);
        painter_.DrawLine(page_x(0, 0), page_y(0, 0), page_x(0, val_len), page_y(0, val_len));
        painter_.DrawLine(page_x(0, zero), page_y(0, zero), page_x(cat_len, zero), page_y(cat_len, zero));
        if (rc.has_secondary && !horiz) painter_.DrawLine(plot_x + plot_w, plot_y, plot_x + plot_w, plot_y + plot_h);

        // Data labels, drawn over everything once the plot is done. Values
        // shown are the series' own (not a 100 % stack's shares).
        std::vector<PendingLabel> labels;
        auto label_text = [&](std::size_t si, std::size_t i) {
            const ChartSeriesData& s = data.series[si];
            const double raw = i < s.values.size() ? s.values[i] : 0.0;
            return data_label_text(labels_at(s, i), s, i < ncat ? data.categories[i] : std::string(), raw);
        };
        // A bar from `base` to `end` (along the value axis) centred at `c`:
        // outside/inside its end, centred, or inside at its base.
        auto bar_label = [&](std::size_t si, std::size_t i, double c, double base, double end, bool negative,
                             const char* default_pos) {
            const DataLabels& dl = labels_at(data.series[si], i);
            if (!dl.shown()) return;
            const std::string pos = dl.position.empty() ? default_pos : dl.position;
            const double inward = (negative ? 1.0 : -1.0) * dl.size * 0.8;  // from the end towards the base
            double v = (base + end) / 2.0;
            LabelPlace place = LabelPlace::Center;
            if (pos == "outEnd") {
                v = end;
                place = horiz ? (negative ? LabelPlace::Left : LabelPlace::Right)
                              : (negative ? LabelPlace::Below : LabelPlace::Above);
            } else if (pos == "inEnd") {
                v = horiz ? end : end + inward;
                if (horiz) place = negative ? LabelPlace::Right : LabelPlace::Left;
            } else if (pos == "inBase") {
                v = horiz ? base : base - inward;
                if (horiz) place = negative ? LabelPlace::Left : LabelPlace::Right;
            }
            labels.push_back({label_text(si, i), page_x(c, v), page_y(c, v), place, &dl});
        };

        // Areas.
        for (const ChartGroup& g : data.groups) {
            if (g.kind != ChartKind::Area || ncat == 0) continue;
            const AxisScale& sc = g.secondary ? secondary.scale : primary.scale;
            std::vector<double> base(ncat, 0.0);
            for (std::size_t si : g.series) {
                std::vector<double> top_v(ncat);
                for (std::size_t i = 0; i < ncat; ++i) {
                    const double v = group_value(data, g, si, i);
                    top_v[i] = g.grouping == Grouping::Clustered ? v : base[i] + v;
                }
                PoDoFo::PdfPainterPath path;
                for (std::size_t i = 0; i < ncat; ++i) {
                    const double c = cat_center(i), v = vpos(sc, top_v[i]);
                    if (i == 0) path.MoveTo(page_x(c, v), page_y(c, v));
                    else        path.AddLineTo(page_x(c, v), page_y(c, v));
                }
                for (std::size_t k = ncat; k-- > 0;) {
                    const double c = cat_center(k);
                    const double v = vpos(sc, g.grouping == Grouping::Clustered ? 0.0 : base[k]);
                    path.AddLineTo(page_x(c, v), page_y(c, v));
                }
                path.Close();
                painter_.GraphicsState.SetNonStrokingColor(pdf_color(data.series[si].color));
                painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Fill);
                for (std::size_t i = 0; i < ncat; ++i) {
                    const DataLabels& dl = labels_at(data.series[si], i);
                    if (!dl.shown()) continue;
                    const double c = cat_center(i), v = vpos(sc, top_v[i]);
                    labels.push_back({label_text(si, i), page_x(c, v), page_y(c, v),
                                      point_label_place(dl.position, LabelPlace::Center), &dl});
                }
                if (g.grouping != Grouping::Clustered) base = top_v;
            }
        }

        // Bars: clustered series side by side, a stacked group as one bar;
        // several bar groups (combo) share the category slot.
        std::size_t units = 0;
        double gap = 1.5;
        bool first_bar_group = true;
        for (const ChartGroup& g : data.groups) {
            if (g.kind != ChartKind::Bar) continue;
            units += g.grouping == Grouping::Clustered ? g.series.size() : 1;
            if (first_bar_group) { gap = g.gap; first_bar_group = false; }
        }
        if (units > 0) {
            const double unit_w = slot / (static_cast<double>(units) + gap);
            std::size_t unit = 0;
            for (const ChartGroup& g : data.groups) {
                if (g.kind != ChartKind::Bar) continue;
                const AxisScale& sc = g.secondary ? secondary.scale : primary.scale;
                const double z = vpos(sc, 0.0);
                if (g.grouping == Grouping::Clustered) {
                    for (std::size_t si : g.series) {
                        painter_.GraphicsState.SetNonStrokingColor(pdf_color(data.series[si].color));
                        for (std::size_t i = 0; i < ncat; ++i) {
                            const double c0 = cat_center(i) - slot / 2.0 + gap * unit_w / 2.0 + static_cast<double>(unit) * unit_w;
                            const double v = group_value(data, g, si, i);
                            fill_rect(c0, c0 + unit_w * 0.95, z, vpos(sc, v));
                            bar_label(si, i, c0 + unit_w * 0.475, z, vpos(sc, v), v < 0.0, "outEnd");
                        }
                        ++unit;
                    }
                } else {
                    for (std::size_t i = 0; i < ncat; ++i) {
                        const double c0 = cat_center(i) - slot / 2.0 + gap * unit_w / 2.0 + static_cast<double>(unit) * unit_w;
                        double pos = 0.0, neg = 0.0;
                        for (std::size_t si : g.series) {
                            const double v = group_value(data, g, si, i);
                            double& acc = v >= 0.0 ? pos : neg;
                            painter_.GraphicsState.SetNonStrokingColor(pdf_color(data.series[si].color));
                            fill_rect(c0, c0 + unit_w * 0.95, vpos(sc, acc), vpos(sc, acc + v));
                            bar_label(si, i, c0 + unit_w * 0.475, vpos(sc, acc), vpos(sc, acc + v), v < 0.0, "ctr");
                            acc += v;
                        }
                    }
                    ++unit;
                }
            }
        }

        // Lines (with markers).
        for (const ChartGroup& g : data.groups) {
            if (g.kind != ChartKind::Line || ncat == 0) continue;
            const AxisScale& sc = g.secondary ? secondary.scale : primary.scale;
            std::vector<double> base(ncat, 0.0);
            for (std::size_t si : g.series) {
                PoDoFo::PdfPainterPath path;
                std::vector<std::pair<double, double>> points;
                for (std::size_t i = 0; i < ncat; ++i) {
                    const double v = group_value(data, g, si, i);
                    const double top_v = g.grouping == Grouping::Clustered ? v : base[i] + v;
                    if (g.grouping != Grouping::Clustered) base[i] = top_v;
                    const double c = cat_center(i), p = vpos(sc, top_v);
                    points.emplace_back(page_x(c, p), page_y(c, p));
                    if (i == 0) path.MoveTo(points.back().first, points.back().second);
                    else        path.AddLineTo(points.back().first, points.back().second);
                }
                const PoDoFo::PdfColor color = pdf_color(data.series[si].color);
                painter_.GraphicsState.SetStrokingColor(color);
                painter_.GraphicsState.SetLineWidth(1.75);
                painter_.DrawPath(path, PoDoFo::PdfPathDrawMode::Stroke);
                if (data.series[si].marker) {
                    painter_.GraphicsState.SetNonStrokingColor(color);
                    for (const auto& [px, py] : points) draw_marker(px, py);
                }
                for (std::size_t i = 0; i < points.size(); ++i) {
                    const DataLabels& dl = labels_at(data.series[si], i);
                    if (!dl.shown()) continue;
                    labels.push_back({label_text(si, i), points[i].first, points[i].second,
                                      point_label_place(dl.position, LabelPlace::Right), &dl});
                }
            }
        }

        draw_data_labels(labels);

        // Reset to black so whatever text/lines follow don't inherit a
        // series color.
        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
        painter_.GraphicsState.SetLineWidth(1.0);
    }

    // A data label waiting to be drawn over the plot: `place` says where it
    // goes relative to (x, y) — above, below, left, right or centred.
    enum class LabelPlace { Above, Below, Left, Right, Center };
    struct PendingLabel {
        std::string       text;
        double            x = 0.0, y = 0.0;
        LabelPlace        place = LabelPlace::Center;
        const DataLabels* style = nullptr;
    };

    void draw_data_labels(const std::vector<PendingLabel>& labels) {
        for (const PendingLabel& l : labels) {
            if (l.text.empty()) continue;
            PoDoFo::PdfFont& font = l.style->bold ? bold_ : regular_;
            const double size = l.style->size;
            const double w = label_width(font, size, l.text);
            const double mid = size * 0.35;  // from the baseline to the middle of the digits
            double x = l.x - w / 2.0, y = l.y - mid;
            switch (l.place) {
                case LabelPlace::Above:  y = l.y + 2.0; break;
                case LabelPlace::Below:  y = l.y - 2.0 - size * 0.75; break;
                case LabelPlace::Left:   x = l.x - w - 3.0; break;
                case LabelPlace::Right:  x = l.x + 3.0; break;
                case LabelPlace::Center: break;
            }
            painter_.GraphicsState.SetNonStrokingColor(pdf_color(l.style->color));
            draw_line(font, size, l.text, x, y);
        }
        painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
    }

    // Where a line/area chart point's label goes (<c:dLblPos>: t, b, l, r, ctr).
    static LabelPlace point_label_place(const std::string& pos, LabelPlace fallback) {
        if (pos == "t")   return LabelPlace::Above;
        if (pos == "b")   return LabelPlace::Below;
        if (pos == "l")   return LabelPlace::Left;
        if (pos == "r")   return LabelPlace::Right;
        if (pos == "ctr") return LabelPlace::Center;
        return fallback;
    }

    void draw_marker(double px, double py) {
        constexpr double r = 2.5;
        PoDoFo::PdfPainterPath diamond;
        diamond.MoveTo(px, py + r);
        diamond.AddLineTo(px + r, py);
        diamond.AddLineTo(px, py - r);
        diamond.AddLineTo(px - r, py);
        diamond.Close();
        painter_.DrawPath(diamond, PoDoFo::PdfPathDrawMode::Fill);
    }

    // Single-row legend: swatch (a line with its marker for line series) +
    // series name, left to right. Series that don't fit in one row are
    // left out rather than wrapped.
    void draw_chart_legend(const ChartData& data, double x, double y) {
        constexpr double kSwatch   = 8.0;
        constexpr double kGap      = 6.0;
        constexpr double kFontSize = 8.0;

        double       cursor_x = x;
        const double max_x    = x + content_width();
        for (std::size_t i = 0; i < data.series.size(); ++i) {
            const ChartSeriesData& ser = data.series[i];
            const bool   line   = ser.kind == ChartKind::Line;
            const double swatch = line ? kSwatch * 2.0 : kSwatch;
            const double text_w = label_width(regular_, kFontSize, ser.name);
            const double item_w = swatch + 4.0 + text_w + kGap * 2.0;
            if (cursor_x + item_w > max_x && cursor_x > x) break;

            if (line) {
                painter_.GraphicsState.SetStrokingColor(pdf_color(ser.color));
                painter_.GraphicsState.SetLineWidth(1.75);
                painter_.DrawLine(cursor_x, y + kSwatch / 2.0, cursor_x + swatch, y + kSwatch / 2.0);
                painter_.GraphicsState.SetNonStrokingColor(pdf_color(ser.color));
                if (ser.marker) draw_marker(cursor_x + swatch / 2.0, y + kSwatch / 2.0);
                painter_.GraphicsState.SetStrokingColor(PoDoFo::PdfColor(0, 0, 0));
                painter_.GraphicsState.SetLineWidth(1.0);
            } else {
                painter_.GraphicsState.SetNonStrokingColor(pdf_color(ser.color));
                painter_.DrawRectangle(cursor_x, y, kSwatch, kSwatch, PoDoFo::PdfPathDrawMode::Fill);
            }
            painter_.GraphicsState.SetNonStrokingColor(PoDoFo::PdfColor(0, 0, 0));
            draw_line(regular_, kFontSize, ser.name, cursor_x + swatch + 4.0, y + 1.0);
            cursor_x += item_w;
        }
    }

    struct PageInfo {
        std::size_t section = 0;
        int         number  = 0;      // displayed page number
        bool        first_in_section = false;
    };

    PoDoFo::PdfMemDocument&   doc_;
    std::vector<Section>&     sections_;
    PageMetrics               pm_;       // current section's (the page being laid out)
    FontBook&                 fonts_;
    PoDoFo::PdfFont&          regular_;   // document default font: chart text
    PoDoFo::PdfFont&          bold_;
    const pugi::xml_document& styles_doc_;
    const StyleSheet&         styles_;
    Numbering&                numbering_;
    const Theme&              theme_;
    double                    default_tab_;
    PoDoFo::PdfPainter        painter_;
    FlowState                 flow_;
    double                    cursor_y_      = 0.0;
    double                    page_top_      = 0.0;  // body top of the current page
    double                    bottom_limit_  = 0.0;  // body bottom of the current page
    std::vector<PageInfo>     pages_;
    std::size_t               section_       = 0;
    int                       next_number_   = 1;
    bool                      first_in_section_ = true;
    PageFields                fields_;               // set while drawing headers/footers
    bool                      page_open_     = false;
    bool                      dry_run_       = false;  // measure only, no painter calls
    bool                      no_page_break_ = false;  // headers/footers: fixed region
    const CellStyle*          cell_style_    = nullptr;  // table style of the cell being laid out
    const LayoutFacts*        known_         = nullptr;  // previous pass, if any
    std::map<std::string, PageNumber> bookmark_pages_;    // this pass
    // Floating objects: drawn over the current page at its end; behindDoc
    // ones found in this pass; the page they're placed on; text-free bands
    // (wrapTopAndBottom, ...) of the current page (PDF y, top > bottom).
    std::vector<FloatingObject> front_floats_;
    std::vector<FloatingObject> behind_floats_;
    std::size_t               float_page_    = 0;
    struct Band { double top = 0.0, bottom = 0.0; };
    std::vector<Band>         bands_;
    bool                      in_cell_       = false;  // anchored drawings stay inline in cells
};

// Whether the body has a field that needs a finished layout: PAGEREF
// (TOC entries, "see page N"), NUMPAGES or SECTIONPAGES.
bool body_needs_layout_facts(pugi::xml_node body) {
    auto needs = [](std::string_view instr) {
        const std::string kw = field_keyword(instr);
        return kw == "PAGEREF" || kw == "NUMPAGES" || kw == "SECTIONPAGES";
    };
    for (pugi::xpath_node n : body.select_nodes(".//w:instrText")) {
        if (needs(n.node().text().get())) return true;
    }
    for (pugi::xpath_node n : body.select_nodes(".//w:fldSimple")) {
        if (needs(n.node().attribute("w:instr").value())) return true;
    }
    return false;
}

// Whether `root` has a floating object behind the text (drawn first on
// its page, which needs a pass that already knows where it goes).
bool has_behind_floats(pugi::xml_node root) {
    for (pugi::xpath_node n : root.select_nodes(".//wp:anchor")) {
        const std::string_view v = n.node().attribute("behindDoc").value();
        if (v == "1" || v == "true" || v == "on") return true;
    }
    return false;
}

} // namespace

#endif  // DOCWEFT_HAVE_PODOFO

void render_native_pdf(const pugi::xml_document& document,
                        const std::unordered_map<std::string, std::string>& parts,
                        const std::filesystem::path& output_path) {
#if !defined(DOCWEFT_HAVE_PODOFO)
    (void)document;
    (void)parts;
    (void)output_path;
    throw ReportException(
        ReportError::NotImplemented,
        "native PDF backend not compiled in — rebuild with "
        "-DDOCWEFT_ENABLE_NATIVE_PDF=ON (requires PoDoFo)");
#else
    using namespace PoDoFo;

    pugi::xml_node body = document.child("w:document").child("w:body");
    if (!body) {
        throw ReportException(
            ReportError::CantCopyDocxTemplate,
            "native PDF backend: word/document.xml has no <w:body>");
    }

    const RelMap rels = build_rel_map(parts);

    // Optional parts: an absent or unparsable one just leaves its document
    // empty (no styles → built-in defaults, no numbering → no list markers).
    auto load_part = [&](const std::string& name, pugi::xml_document& out) {
        if (const auto it = parts.find(name); it != parts.end()) {
            out.load_buffer(it->second.data(), it->second.size());
        }
    };
    pugi::xml_document styles_doc;
    pugi::xml_document numbering_doc;
    pugi::xml_document settings_doc;
    load_part("word/styles.xml", styles_doc);
    load_part("word/numbering.xml", numbering_doc);
    load_part("word/settings.xml", settings_doc);

    Theme theme;
    for (const auto& [name, bytes] : parts) {
        if (name.rfind("word/theme/", 0) == 0 && name.find("/_rels/") == std::string::npos) {
            pugi::xml_document theme_doc;
            if (theme_doc.load_buffer(bytes.data(), bytes.size())) theme.load(theme_doc);
            break;
        }
    }
    const StyleSheet styles(styles_doc, theme);
    Numbering numbering(numbering_doc);
    const double default_tab = twips_to_pt(
        settings_doc.child("w:settings").child("w:defaultTabStop").attribute("w:val").as_llong(720));

    // Sections: each ends at a paragraph carrying <w:pPr>/<w:sectPr>; the
    // body's own <w:sectPr> describes the last one.
    std::vector<Section> sections(1);
    std::vector<pugi::xml_node> sect_prs;  // one per section, in order
    for (pugi::xml_node b : blocks_of(body)) {
        sections.back().blocks.push_back(b);
        if (pugi::xml_node sp = b.child("w:pPr").child("w:sectPr")) {
            sect_prs.push_back(sp);
            sections.emplace_back();
        }
    }
    sect_prs.push_back(body.child("w:sectPr"));

    std::vector<std::unique_ptr<pugi::xml_document>> hf_docs;
    const bool even_and_odd = on_off(settings_doc.child("w:settings").child("w:evenAndOddHeaders"));
    for (std::size_t si = 0; si < sections.size(); ++si) {
        Section& sec = sections[si];
        pugi::xml_node sect = sect_prs[si];
        // Header/footer kinds a section doesn't define carry over from the
        // previous section, as in Word.
        if (si > 0) sec.hf = sections[si - 1].hf;
        sec.pm            = read_page_metrics(sect);
        sec.hf.title_page = on_off(sect.child("w:titlePg"));
        sec.hf.even_and_odd = even_and_odd;
        const std::string_view type = sect.child("w:type").attribute("w:val").value();
        sec.start = type == "continuous" ? Section::Start::Continuous
                  : type == "evenPage"   ? Section::Start::EvenPage
                  : type == "oddPage"    ? Section::Start::OddPage
                                         : Section::Start::NextPage;
        sec.page_number_start = sect.child("w:pgNumType").attribute("w:start").as_int(0);
        sec.page_number_format = sect.child("w:pgNumType").attribute("w:fmt").value();

        auto load_hf = [&](const char* ref_tag, std::array<HeaderFooterPart, 3>& slots) {
            for (pugi::xml_node ref : sect.children(ref_tag)) {
                const std::string_view kind = ref.attribute("w:type").value();
                const int slot = kind == "first" ? HeaderFooterSet::First
                               : kind == "even"  ? HeaderFooterSet::Even
                                                 : HeaderFooterSet::Default;
                const auto rel = rels.find(ref.attribute("r:id").value());
                if (rel == rels.end()) continue;
                const std::string part_name = "word/" + rel->second;
                const auto part = parts.find(part_name);
                if (part == parts.end()) continue;
                auto doc = std::make_unique<pugi::xml_document>();
                if (!doc->load_buffer(part->second.data(), part->second.size())) continue;

                const std::string file = std::filesystem::path(part_name).filename().string();
                HeaderFooterPart& hfp = slots[static_cast<std::size_t>(slot)];
                hfp.root    = doc->first_child();
                hfp.blocks  = blocks_of(hfp.root);
                hfp.rels    = build_rel_map(parts, "word/_rels/" + file + ".rels");
                hfp.present = true;
                hf_docs.push_back(std::move(doc));
            }
        };
        load_hf("w:headerReference", sec.hf.header);
        load_hf("w:footerReference", sec.hf.footer);
    }

    if (strict_mode_enabled()) {
        UnsupportedScan scan(parts, styles, numbering);
        scan.scan(body, rels);
        std::set<const void*> scanned;
        for (const Section& sec : sections) {
            for (const auto* set : {&sec.hf.header, &sec.hf.footer}) {
                for (const HeaderFooterPart& part : *set) {
                    if (part.present && scanned.insert(part.root.internal_object()).second) {
                        scan.scan(part.root, part.rels);
                    }
                }
            }
        }
        if (!scan.found().empty()) {
            std::string list;
            for (const std::string& item : scan.found()) {
                if (!list.empty()) list += "; ";
                list += item;
            }
            throw ReportException(
                ReportError::NotImplemented,
                fmt::format("native PDF backend can't render this document faithfully — "
                            "the PDF would come out incomplete or misleading. Not supported: "
                            "{}. Set DOCWEFT_NATIVE_PDF_STRICT=0 to render it anyway.",
                            list));
        }
    }

    try {
        const TextStyle default_style = styles.mark_style(pugi::xml_node{});
        // One pass, into its own document (list counters start over).
        auto render = [&](const LayoutFacts* known, LayoutFacts& learned) {
            auto doc = std::make_unique<PdfMemDocument>();
            Numbering pass_numbering(numbering_doc);
            FontBook fonts(*doc, default_style.font);
            Layout layout(*doc, sections, fonts, default_style, styles_doc, styles, pass_numbering, theme,
                          default_tab, known);
            layout.measure_headers_footers(parts);
            layout.layout_sections(rels, parts);
            layout.finish();
            layout.draw_headers_footers(parts);
            learned = layout.facts();
            return doc;
        };

        // Fields in the body that need the finished layout (TOC page
        // numbers, cross-references, "page X of Y" in body text) and
        // pictures behind the text: lay out again with what the previous
        // pass found, until nothing moves —
        // real numbers can be wider than the cached ones — at most 3 passes.
        LayoutFacts facts;
        std::unique_ptr<PdfMemDocument> doc = render(nullptr, facts);
        bool more_passes = body_needs_layout_facts(body) || has_behind_floats(body);
        for (const Section& sec : sections) {
            for (const auto* set : {&sec.hf.header, &sec.hf.footer}) {
                for (const HeaderFooterPart& part : *set) {
                    if (part.present && has_behind_floats(part.root)) more_passes = true;
                }
            }
        }
        if (more_passes) {
            for (int pass = 2; pass <= 3; ++pass) {
                LayoutFacts next;
                doc = render(&facts, next);
                const bool settled = next == facts;
                facts = std::move(next);
                if (settled) break;
            }
        }
        doc->Save(output_path.string());
    } catch (const ReportException&) {
        throw;
    } catch (const PdfError& e) {
        throw ReportException(
            ReportError::SaveFailed,
            fmt::format("native PDF backend: PoDoFo failed: {}", e.what()));
    }
#endif
}

} // namespace docweft::docx
