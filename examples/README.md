# TextFabric examples

Minimal binaries that show how the library is meant to be consumed.

| Binary | Purpose | Library used |
|---|---|---|
| `generate_template` | Produces a valid OOXML `template.docx` with 4 header bookmarks + a 2-column table with a template `<w:tr>` | libzip (no textfabric) |
| `basic_report` | Loads a template, fills header bookmarks, clones a table row per data entry, saves `.docx` + (if LibreOffice is available) `.pdf` + `.html` | `textfabric` public API only |
| `native_pdf_report` | Builds its own chart-free template, fills it, and saves straight to `.pdf` via the native PoDoFo backend with Microsoft Word/LibreOffice detection explicitly disabled — proves the PDF really was produced without either. Only built when `-DTEXTFABRIC_ENABLE_NATIVE_PDF=ON` (see [PLAN.md](../PLAN.md)); a sample output sits alongside this file at [`native_report.pdf`](native_report.pdf). | `textfabric` public API + libzip (to build its own template) |

## Build

From the repository root:

```bash
cmake --preset linux-x64 -DTEXTFABRIC_BUILD_EXAMPLES=ON
cmake --build --preset linux-x64
```

or (without presets):

```bash
cmake -B build -DTEXTFABRIC_BUILD_EXAMPLES=ON -DTEXTFABRIC_USE_FETCHCONTENT=ON
cmake --build build -j
```

Both binaries, the generated `template.docx` and a copy of the sample `logo.png` (75 bytes, committed at `examples/logo.png`) land in `build/examples/`.

## Run

```bash
cd build/examples
./basic_report                              # uses ./template.docx → ./report.docx
./basic_report template.docx custom.docx    # explicit paths
```

Expected output (with LibreOffice installed):

```
TextFabric v0.1.0
  template: template.docx
  output:   report.docx

Wrote report.docx
Open it in MS Word / LibreOffice to verify:
  - title 'Q1 2026 Summary'
  - author 'Иван Петров'
  - today's date (2026-04-22)
  - greeting 'Hello Alice, welcome to Acme Corp.'
  - Scores table with 3 rows:
      Alice — 95
      Борис — 82
      Carol Жу — 77
  - embedded PNG logo from logo.png (via setImage, upscaled to 64×64 via ImageSize{min=64, max=256})
Wrote report.pdf (PDF via Word or LibreOffice, whichever was found)
Wrote report.html (HTML via Word or LibreOffice, whichever was found)
```

Without LibreOffice (or with `TEXTFABRIC_SOFFICE=""`) the last two lines become:

```
PDF: skipped — LibreOffice not found. Install soffice or set TEXTFABRIC_SOFFICE to enable PDF/HTML export.
HTML: skipped — LibreOffice not found. Install soffice or set TEXTFABRIC_SOFFICE to enable PDF/HTML export.
```

The `.docx` is still written in both cases — PDF/HTML are derived outputs that never block the primary save path.

Open `report.docx` in Word / LibreOffice / WPS — the placeholders must be
replaced with runtime values while all styles (bold, fonts, sizes) survive
intact.

## What the template looks like

`generate_template` writes a minimal OOXML package containing:

```
template.docx/
├── [Content_Types].xml
├── _rels/.rels
├── word/
│   ├── _rels/document.xml.rels
│   ├── document.xml         ← 4 bookmarks + placeholders
│   └── styles.xml
```

Bookmarks and placeholders inside `word/document.xml` (the `Body.Logo` picture is injected at runtime by `basic_report`, not baked into `template.docx`):

| Bookmark             | Placeholder(s)              | What the report fills in |
|----------------------|-----------------------------|--------------------------|
| `_Header.ReportName` | `{{title}}`                 | "Q1 2026 Summary"        |
| `_Header.User`       | `{{user}}`                  | "Иван Петров"            |
| `_Header.UKDate`     | `{{date}}`                  | today's date             |
| `Body.Greeting`      | `{{name}}`, `{{app}}`       | "Alice", "Acme Corp" |
| `Body.Scores` (inside `<w:tr>`) | `{{rowName}}`, `{{rowScore}}` | cloned per data row via `setTableRow` — Alice/95, Борис/82, Carol Жу/77 |
| `Body.Logo`          | `{{logo}}`                  | committed `logo.png` embedded via `setImage(…, ImageSize{64,64,256,256})` → 2×2 PNG upscaled to 64×64 px displayed, aspect 1:1 preserved |

## Integrating in your own CMake project

```cmake
find_package(TextFabric 0.1 CONFIG REQUIRED)

add_executable(my_app main.cpp)
target_link_libraries(my_app PRIVATE TextFabric::textfabric)
```

```cpp
#include <textfabric/merger.hpp>
#include <textfabric/error.hpp>

auto merger = textfabric::make_docx_merger();
merger->load("template.docx");

// Header substitution — single value per bookmark/field.
merger->setClipboardValue("_Header.User", "{{user}}", "Alice");

// Table row cloning — template <w:tr> with {{rowName}}/{{rowScore}} placeholders
// becomes three filled rows, template row itself is removed.
merger->setTableRow("Body.Scores",
                    {"{{rowName}}", "{{rowScore}}"},
                    {
                        {"Alice", "95"},
                        {"Bob",   "82"},
                        {"Carol", "77"},
                    });

merger->save("report.docx");
```
