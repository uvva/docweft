#pragma once

#include <pugixml.hpp>

#include <filesystem>
#include <string>
#include <unordered_map>

namespace docweft::docx {

/// Renders the loaded document straight to PDF via PoDoFo, without shelling
/// out to Microsoft Word or LibreOffice. Scoped to what DocWeft itself
/// understands: paragraphs/runs (font size + bold, first run's style wins
/// per paragraph — same convention as the rest of the merger), simple grid
/// tables, embedded PNG images, and bar/line/area/pie/doughnut charts (3D
/// variants rendered flat, no projection). Anything else — pie-of-pie/
/// bar-of-pie, radar, scatter, bubble, stock, and surface charts, most
/// prominently — throws ReportException{NotImplemented}. See PLAN.md
/// ("Functional / feature-completeness risk") for the full list of known
/// gaps.
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
///                     uses a feature this renderer doesn't cover (a
///                     chart type outside bar/line/area/pie/doughnut,
///                     most prominently).
///   - SaveFailed:     PoDoFo itself failed to write the PDF (caught
///                     PdfError is wrapped with its message).
void render_native_pdf(const pugi::xml_document& document,
                        const std::unordered_map<std::string, std::string>& parts,
                        const std::filesystem::path& output_path);

} // namespace docweft::docx
