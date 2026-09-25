<p align="center">
  <img src="assets/logo.png" alt="DocWeft logo" width="140">
</p>

<h1 align="center">DocWeft</h1>
<p align="center"><b>Fill Word templates by bookmark name — no template engine, no intermediate format.</b></p>

<p align="center">
  <a href="LICENSE"><img alt="License: MIT" src="https://img.shields.io/badge/license-MIT-blue.svg"></a>
  <img alt="C++ standard" src="https://img.shields.io/badge/C%2B%2B-20-00599C?logo=c%2B%2B&logoColor=white">
  <img alt="Platforms" src="https://img.shields.io/badge/platform-Linux%20%7C%20Windows%20%7C%20macOS-lightgrey">
  <img alt="Tests" src="https://img.shields.io/badge/tests-69%2F69%20passing-brightgreen">
</p>

Cross-platform C++20 library for filling DOCX templates: text substitution, table row cloning, embedded chart updates, and image insertion — all addressed by named Word bookmarks, with no intermediate format. Optionally converts the result to PDF/HTML via Microsoft Word or LibreOffice.

**Confirmed platforms (v0.1.0):** Linux x86_64 (GCC + Nix) — `ctest` **69/69 ✅**, covering charts, image handling (PNG/JPEG/BMP/TIFF), the PDF/HTML converters, and 3 integration tests against `find_package(DocWeft)` from an install prefix. Windows x64 (MSVC 2022 + vcpkg) — builds and passes the core suite; the JPEG/BMP/TIFF branches are still waiting on a dedicated Windows-host run. Windows x86 and macOS — CMake presets are ready, untested.

## Features

- 🔖 **Bookmark-addressed substitution** — target any named Word bookmark; no custom templating syntax beyond a plain `{{token}}` or `kNAME` placeholder
- 📊 **Table row cloning** — one template `<w:tr>` in, N filled rows out
- 📈 **Chart updates** — rewrite data points in embedded bar/line/pie/area/doughnut/radar/3D charts
- 🖼️ **Image insertion** — PNG passes through untouched, JPEG/BMP/TIFF are transcoded on the fly, with aspect-ratio-preserving size bounds
- 📄 **PDF/HTML export** — routed through Microsoft Word or LibreOffice, no conversion pipeline to wire up yourself
- 🌍 **UTF-8 throughout** — Cyrillic, CJK, whatever your template contains (the Cyrillic names in the example below aren't a typo — they're there to prove it)

## Table of Contents

- [Quick Start](#quick-start)
- [Authoring a `.docx` Template](#authoring-a-docx-template)
- [API](#api)
- [Full Example](#full-example)
- [Runtime Dependencies](#runtime-dependencies)
- [Error Codes](#error-codes)
- [Build Dependencies](#build-dependencies)
- [License](#license)

---

## Quick Start

Building from the repository:

```bash
# Linux (Nix dev shell):
nix develop
cmake --preset linux-x64
cmake --build --preset linux-x64
ctest --preset linux-x64

# Windows (MSVC 2022 + vcpkg) — full walkthrough in docs/build-windows.md:
cmake --preset win-x64-vcpkg
cmake --build --preset win-x64-vcpkg
ctest   --preset win-x64-vcpkg
```

Consuming from another CMake project:

```cmake
find_package(DocWeft 0.1 CONFIG REQUIRED)
target_link_libraries(my_app PRIVATE DocWeft::docweft)
```

Or via a vcpkg manifest (`vcpkg.json`):

```json
{ "dependencies": ["docweft"] }
```

Minimal call:

```cpp
#include <docweft/merger.hpp>

auto m = docweft::make_docx_merger();
m->setCodePage("UTF-8");
m->load("template.docx");
m->setClipboardValue("Body.Greeting", "{{name}}", "Alice");
m->save("report.docx");
```

---

## Authoring a `.docx` Template

The template is a plain DOCX authored in MS Word or LibreOffice Writer — no intermediate format required.

### 1. Bookmarks (`<w:bookmarkStart w:name="..."/>`)

Every request is addressed **by Word bookmark name**. Drop a bookmark over the text that should be replaced (Word: *Insert → Bookmark → Add*).

| Bookmark type | Purpose | Example |
|---|---|---|
| `_Header` / `_Footer` | **Pseudo-bookmarks**, reserved for `word/header*.xml` / `word/footer*.xml`. No real `<w:bookmarkStart>` is required or looked up — the engine walks every header/footer part in the archive and substitutes `value` into the first one where `field` occurs. Handy when the template's header/footer holds bare text placeholders with no bookmarks. | `setClipboardValue("_Header", "{User}", "Alice")` |
| Regular bookmark | Any name inside `word/document.xml`, e.g. `Body.*` | `Body.Greeting`, `Body.Scores` |

### 2. Placeholders

Inside a bookmark, place tokens like `{{name}}` or `kNAME`. The library finds the **first occurrence** of the token in the bookmark's concatenated text and replaces it; the first run's formatting (`<w:rPr>`: font, bold, size) is preserved.

A token may be split across several `<w:r>`/`<w:t>` elements (a common artifact of editing in Word) — the engine correctly re-joins the text and cleans up the trailing `<w:t>` fragments.

### 3. Tables

For dynamically cloned rows:

1. Draw a table: one **header** row (left untouched) + one **template** row.
2. Put **a single bookmark** on the template row, spanning it from the first cell to the last.
3. Give each cell a **unique placeholder** (`{{col1}}`, `{{col2}}`, …).

Example fragment of `word/document.xml`:

```xml
<w:tr>  <!-- template row -->
  <w:tc><w:p>
    <w:bookmarkStart w:id="5" w:name="Body.Scores"/>
    <w:r><w:t>{{rowName}}</w:t></w:r>
  </w:p></w:tc>
  <w:tc><w:p>
    <w:r><w:t>{{rowScore}}</w:t></w:r>
    <w:bookmarkEnd w:id="5"/>
  </w:p></w:tc>
</w:tr>
```

v1 limitations:
- Nested tables haven't been tested, but the cloning algorithm doesn't break them.
- `<w:bookmarkStart>` elements are stripped from clones automatically to avoid duplicate `w:id`s (Word would otherwise refuse to open the document).

### 4. Images (`setImage`)

The bookmark (e.g. `Body.Logo`) must sit **inside a single `<w:p>`**; `setImage` removes every `<w:r>` between `bookmarkStart`/`bookmarkEnd` and inserts `<w:r><w:drawing>…</w:drawing></w:r>` in their place. A placeholder like `{{logo}}` is replaced by the image wholesale.

The on-page size (`<wp:extent cx cy>`) defaults to the source file's pixel dimensions at 96 DPI (9525 EMU/px). The `ImageSize{min_width_px, min_height_px, max_width_px, max_height_px}` parameter lets you set bounds — a tiny icon is upscaled to `min`, an oversized photo is downscaled to `max`, aspect ratio is preserved via a single scalar applied to both axes. Conflicting bounds (e.g. `min_width=400, max_height=200` on a square image) → `ReportError::InvalidField` **before** the archive is mutated, so the document is left untouched.

---

## API

Full header: [include/docweft/merger.hpp](include/docweft/merger.hpp). The factory `docweft::make_docx_merger()` returns `std::unique_ptr<IReportMerger>`. All strings are UTF-8.

| Method | What it does |
|---|---|
| `setCodePage(cp)` | Sets the encoding. v1 accepts only `"UTF-8"`; anything else → `NotImplemented`. |
| `load(path)` | Reads the `.docx` archive, parses `word/document.xml`. |
| `save(path.docx)` | Round-trips the `.docx` archive. |
| `save(path.pdf/html)` | Conversion via Microsoft Word, LibreOffice, the optional native PDF renderer or the optional remote converter — see [PDF / HTML export](#pdf--html-export). `NoConverter` if none is available. |
| `setClipboardValue(bookmark, field, value)` | Substitutes `value` for `field` inside `bookmark`. Handles split runs. `bookmark` = `"_Header"`/`"_Footer"` is a special case: it walks every `word/header*.xml`/`word/footer*.xml` part in the archive instead of resolving a bookmark. |
| `hasBookmark(bookmark)` | Checks whether the loaded template has a bookmark with that name. Returns `false` for both a missing bookmark and an unloaded template — never throws. Does not resolve the `_Header`/`_Footer` pseudo-bookmarks. |
| `clearBookmark(bookmark)` | Wipes the placeholder text inside a bookmark, leaving it empty — useful when a bookmark is deliberately left unfilled on a given run. Does not resolve the `_Header`/`_Footer` pseudo-bookmarks. |
| `setTableRow(bookmark, fields, rows)` | Clones the `<w:tr>` template once per entry in `rows`, substitutes `rows[i][j]` into the cell for `fields[j]`, then removes the template row. |
| `paste(bookmark)` | Validates that the bookmark exists and logs the activation. A no-op in v1, kept for a future "expand section" semantic. |
| `setChartValue(bookmark, field, series, category, value)` | Rewrites `<c:val>/<c:numCache>/<c:pt idx="N">/<c:v>` for the `(series, category)` point in the chart the bookmark points at. Supports bar/line/pie/area/doughnut/radar/3D variants. Scatter/bubble/stock/surface → `NotImplemented`. `field` is accepted for API symmetry but unused in matching. **Limitation:** only updates an already-existing `(series, category)` pair — it cannot add or remove a category or series. The chart's category axis and series set are fixed by the template; a `category` absent from the template raises `InvalidField`. |
| `setImage(bookmark, image_path, bounds={})` | Inserts an image into the bookmark's paragraph. PNG passes through as-is; JPEG/BMP go through stb_image; TIFF goes through libtiff. Everything is transcoded to PNG before embedding. The optional `ImageSize{min_w, min_h, max_w, max_h}` (in pixels) clamps the displayed size while preserving aspect ratio. |

---

## Full Example

See working code in [examples/basic_report.cpp](examples/basic_report.cpp) and the template factory in [examples/generate_template.cpp](examples/generate_template.cpp). Abridged skeleton:

```cpp
#include <docweft/error.hpp>
#include <docweft/merger.hpp>

try {
    auto m = docweft::make_docx_merger();
    m->setCodePage("UTF-8");
    m->load("template.docx");

    // --- Header fields (walks word/header*.xml, no bookmark needed) ---
    m->setClipboardValue("_Header", "{ReportName}", "Q1 2026 Summary");
    m->setClipboardValue("_Header", "{User}",       "Иван Петров");
    m->setClipboardValue("_Header", "{UKDate}",     "2026-04-22");

    // --- Mixed-field bookmark ---
    m->setClipboardValue("Body.Greeting", "{{name}}", "Alice");
    m->setClipboardValue("Body.Greeting", "{{app}}",  "Acme Corp");

    // --- Table rows ---
    m->setTableRow("Body.Scores",
        {"{{rowName}}", "{{rowScore}}"},
        {
            {"Alice",    "95"},
            {"Борис",    "82"},
            {"Carol Жу", "77"},
        });

    // --- Image ---
    // No bounds → displayed size = pixel size. With bounds → aspect-
    // preserving clamp (upscales tiny icons, downscales large photos).
    m->setImage("Body.Logo", "logo.png",
                docweft::ImageSize{/*min*/64, 64, /*max*/512, 512});

    m->paste("Body.Greeting");
    m->save("report.docx");
}
catch (const docweft::ReportException& e) {
    std::cerr << docweft::to_string(e.code()) << ": " << e.what() << "\n";
}
```

---

## Runtime Dependencies

`load`, `save(".docx")`, `setClipboardValue`, `setTableRow`, `paste`, and `setImage(PNG)` need no external dependencies — everything lives in the library itself (the PNG parser is hand-rolled).

### Images (`setImage`)

| Input format | What happens | Requires |
|---|---|---|
| **PNG** | Bytes pass straight through, dimensions are read from the IHDR chunk | — (hand-rolled parser in `src/docx/image.cpp`) |
| **JPEG**, **BMP** | Decoded to RGBA → re-encoded as PNG → embedded | `stb_image` + `stb_image_write` (vcpkg `stb` / FetchContent from github/nothings/stb). Without them → `NotImplemented`. |
| **TIFF** | `TIFFClientOpen` → `TIFFReadRGBAImageOriented` → repack → PNG | `libtiff` (vcpkg `tiff` / nix `libtiff`). Without libtiff → `NotImplemented`. |

Format detection is by magic bytes on the input file, not by extension. The archive always stores `word/media/image{N}.png` — nothing but `<Default Extension="png" ContentType="image/png"/>` shows up in `[Content_Types].xml` on output.

### PDF / HTML export

`save("*.pdf")` / `save("*.html")` walks a chain of converters, top to bottom. Every converter that is present is tried in turn; if one fails, the next one gets the document.

| Step | Converter | Formats | How |
|---|---|---|---|
| 1 | **Microsoft Word** (Windows, macOS) | pdf, html | **Windows:** `reg query HKCR\Word.Application`; conversion through a generated `.vbs` + `cscript //B //Nologo` → `Documents.Open` → `SaveAs2`, headless. **macOS:** `Microsoft Word.app` in `/Applications` or `~/Applications`; conversion through AppleScript (`osascript`) → `open` → `save as`, Word is quit afterwards if it wasn't already running. |
| 2 | **LibreOffice `soffice`** (all platforms) | pdf, html | `DOCWEFT_SOFFICE` env var (absolute path; an empty value disables this branch) → Windows `C:\Program Files[(x86)]\LibreOffice\program\soffice.exe` → macOS `[~]/Applications/LibreOffice.app/Contents/MacOS/soffice` → PATH probe (`soffice --version`). Conversion runs `soffice --headless --convert-to pdf/html --outdir <scratch> <input>`. |
| 3 | **Native PDF renderer** (PoDoFo) | pdf | Built in with `-DDOCWEFT_ENABLE_NATIVE_PDF=ON` (default OFF). Best-effort layout, no external program; see [PLAN.md](PLAN.md) for the supported subset. |
| 4 | **Remote converter** | pdf | Built in with `-DDOCWEFT_ENABLE_REMOTE_CONVERTER=ON` (default OFF) and active only when `DOCWEFT_CONVERTER_URL` is set — there is no default endpoint. Sends a multipart POST with the `.docx` in the `files` field and expects PDF bytes back ([Gotenberg](https://gotenberg.dev)'s `/forms/libreoffice/convert` contract), using the `curl` command-line tool. |
| — | nothing available | | `ReportError::NoConverter`; the message lists each converter and why it was skipped. If converters were available but all of them failed — `SaveFailed` with each failure listed (or that converter's own code when only one was tried). |

Environment variables:

| Name | Effect |
|---|---|
| `DOCWEFT_DISABLE_CONVERTERS=1` | Master switch: the whole chain is off, `save("*.pdf")` → `NoConverter`. Used by tests. |
| `DOCWEFT_NO_MSWORD=1` | Skips the Word branch (Windows and macOS). Useful when Word is installed but you need LibreOffice-compatible rendering. |
| `DOCWEFT_SOFFICE=<path>` | If set non-empty — use only this binary as LibreOffice, skipping PATH/Program Files/app bundle lookups. If set empty — disable the LibreOffice branch entirely. |
| `DOCWEFT_SOFFICE_TIMEOUT=<seconds>` | LibreOffice time limit, default 120. On expiry `soffice` and the processes it started are killed and the chain moves on to the next converter. |
| `DOCWEFT_CONVERTER_URL=<url>` | Remote converter endpoint, e.g. `http://localhost:3000/forms/libreoffice/convert`. |
| `DOCWEFT_CONVERTER_TOKEN=<token>` | Optional; sent as `Authorization: Bearer <token>`. Passed to curl through a config file, never on the command line, and never included in error messages. |
| `DOCWEFT_CONVERTER_TIMEOUT=<seconds>` | Remote request timeout, default 120. |

The MSWord branch is a deliberate trade-off: on workstations with Office installed it produces output identical to what the user sees interactively in Word, without requiring LibreOffice alongside it. The cost is a possible rendering difference between machines with and without Word; for scenarios where cross-platform rendering parity matters, set `DOCWEFT_NO_MSWORD=1` and deploy LibreOffice everywhere.

On macOS the first conversion through Word shows the system prompt allowing your app to control Microsoft Word (Automation). If it's denied, the Word step fails with a hint to allow it in System Settings → Privacy & Security → Automation, and the chain moves on to LibreOffice. Your app bundle needs `NSAppleEventsUsageDescription` in `Info.plist`, and under hardened runtime the `com.apple.security.automation.apple-events` entitlement — without them macOS rejects the request silently. Word is sandboxed, so it may also ask once to grant access to the temp folder.

## Error Codes

Every exception is a [`docweft::ReportException`](include/docweft/error.hpp) carrying a `code()` of type `ReportError`:

| Code | Thrown when |
|---|---|
| `CantOpenTemplate`     | The file doesn't exist or isn't a zip archive |
| `CantCopyDocxTemplate` | The archive has no `word/document.xml`, or the XML fails to parse |
| `InvalidBookmark`      | The bookmark wasn't found; or `setTableRow` targets a bookmark outside a `<w:tr>` |
| `InvalidField`         | The placeholder is missing from the bookmark; or `rows[i].size() != fields.size()` |
| `SaveFailed`           | An unrecognized extension was passed to `save()`, or the zip write failed |
| `NoConverter`          | No conversion tool is available for `save("*.pdf")` / `save("*.html")` — see [PDF / HTML export](#pdf--html-export). Fix: install MS Word (Windows/macOS) or LibreOffice, set `DOCWEFT_SOFFICE=/abs/path/to/soffice`, build with the native/remote converter (PDF only), or unset `DOCWEFT_DISABLE_CONVERTERS`. |
| `NotImplemented`       | The scenario isn't supported by the current implementation: `setCodePage(≠ "UTF-8")`; `setChartValue` on a scatter/bubble/stock/surface chart; JPEG/BMP/TIFF in `setImage` when the library was built without stb_image/libtiff. |

`docweft::to_string(ReportError)` returns the code's name — handy for logs and user-facing messages.

---

## Build Dependencies

pugixml, libzip, nlohmann-json, fmt, inja, stb, libtiff — via the vcpkg manifest (Windows/macOS) or the Nix flake (Linux). Details in [vcpkg.json](vcpkg.json) and [flake.nix](flake.nix). Tests use Catch2.

## License

[MIT](LICENSE).
