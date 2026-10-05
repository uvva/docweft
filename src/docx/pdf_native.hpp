#pragma once

#include <pugixml.hpp>

#include <filesystem>
#include <string>
#include <unordered_map>

namespace docweft::docx {

/// Renders the loaded document straight to PDF via PoDoFo, without shelling
/// out to Microsoft Word or LibreOffice. Paragraphs are laid out here:
/// paragraph/character styles, fonts found by name (with metric-compatible
/// substitutes), per-run bold/italic/underline/strike/color/size and
/// (small) capitals, superscript/subscript and raised/lowered text,
/// alignment (incl. justified), paragraph borders and shading, indents,
/// spacing, tab stops (with leaders), list numbering, hidden text
/// (<w:vanish>) left out, fields (PAGE/NUMPAGES/SECTIONPAGES and PAGEREF
/// — TOC page numbers — filled in from its own layout), keep with next,
/// headers/footers and page breaks, several sections (page size,
/// orientation, margins, headers/footers, page numbering restarts and
/// formats); tables with borders, shading, table styles (following
/// <w:basedOn>) and their regions (header row, banding, ...), margins,
/// vertical alignment, row heights, vertically merged cells and repeated
/// header rows; footnotes at the bottom of their reference's page (long
/// ones continued on the next) and endnotes at the end of the document or
/// section, numbered as <w:footnotePr>/<w:endnotePr> say; PNG/JPEG
/// images, inline or floating (<wp:anchor>: positioned on the page or in
/// their table cell, in front of or behind
/// the text, text kept out of their band); text boxes and simple shapes
/// (rectangle, rounded rectangle, ellipse, line) with fill, outline and
/// text; bar/line/area/pie/doughnut charts — stacked, 100 % stacked,
/// horizontal, combined with a secondary axis — with the template's
/// colors, titles and data labels, numbers in the document's language.

///
/// Strict mode (on unless DOCWEFT_NATIVE_PDF_STRICT=0): before drawing, the
/// document is checked for content this renderer would get wrong rather
/// than just plainer — equations, multi-column layout,
/// non-PNG/JPEG images, SmartArt, shape groups, freeform or rotated
/// shapes, legacy VML graphics, radar/scatter/...
/// charts, pie charts mixed with other types, ... — and rejected with
/// NotImplemented listing all of them, so the caller's converter chain can
/// move on.
///
/// `document` is the already-parsed `word/document.xml` DOM (DocxMerger
/// keeps this live and re-serializes it into `parts` on save — both must be
/// in sync, i.e. call this only where the .docx save path would also be
/// valid). `parts` supplies everything else needed to resolve images:
/// `word/_rels/document.xml.rels` for relationship IDs and the raw
/// `word/media/imageN.png` bytes themselves.
///
/// Only meaningfully implemented when built with
/// DOCWEFT_ENABLE_NATIVE_PDF (which defines DOCWEFT_HAVE_PODOFO) —
/// otherwise every call throws ReportException{NotImplemented}, so callers
/// don't need to guard the call site with an #ifdef.
///
/// Throws ReportException:
///   - NotImplemented: PoDoFo support wasn't compiled in, or the document
///                     uses something this renderer can't draw faithfully
///                     (see strict mode above).
///   - SaveFailed:     PoDoFo itself failed to write the PDF (caught
///                     PdfError is wrapped with its message).
void render_native_pdf(const pugi::xml_document& document,
                        const std::unordered_map<std::string, std::string>& parts,
                        const std::filesystem::path& output_path);

} // namespace docweft::docx
