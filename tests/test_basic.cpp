#include <catch2/catch_test_macros.hpp>

#include "docweft/docweft.hpp"
#include "docweft/docx/reader.hpp"
#include "docweft/docx/writer.hpp"
#include "docweft/template/engine.hpp"

#include <filesystem>
#include <fstream>

namespace fs = std::filesystem;

// ── file_exists ────────────────────────────────────────────────────────────

TEST_CASE("file_exists detects existing files", "[core]") {
    // Create a temporary file next to the test binary
    const fs::path tmp = fs::temp_directory_path() / "docweft_test.txt";
    {
        std::ofstream ofs(tmp);
        ofs << "hello from DocWeft";
    }

    REQUIRE(docweft::file_exists(tmp));

    // Clean up
    fs::remove(tmp);
}

TEST_CASE("file_exists returns false for missing files", "[core]") {
    REQUIRE_FALSE(docweft::file_exists("/nonexistent/file.txt"));
}

TEST_CASE("file_exists returns false for directories", "[core]") {
    REQUIRE_FALSE(docweft::file_exists(fs::temp_directory_path()));
}

// ── read_file ──────────────────────────────────────────────────────────────

TEST_CASE("read_file reads content correctly", "[core]") {
    const fs::path tmp = fs::temp_directory_path() / "docweft_read.txt";
    const std::string content = "DocWeft test content\nline two";
    {
        std::ofstream ofs(tmp);
        ofs << content;
    }

    REQUIRE(docweft::read_file(tmp) == content);

    fs::remove(tmp);
}

TEST_CASE("read_file returns empty string for missing file", "[core]") {
    REQUIRE(docweft::read_file("/no/such/file.txt").empty());
}

// ── docx::Reader stub ──────────────────────────────────────────────────────

TEST_CASE("Reader reports not open for missing file", "[docx]") {
    docweft::docx::Reader reader("/nonexistent.docx");
    REQUIRE_FALSE(reader.is_open());
    REQUIRE(reader.document_xml().empty());
}

TEST_CASE("Reader reports open for existing file", "[docx]") {
    const fs::path tmp = fs::temp_directory_path() / "fake.docx";
    { std::ofstream ofs(tmp); ofs << "PK"; } // minimal content

    docweft::docx::Reader reader(tmp);
    REQUIRE(reader.is_open());

    fs::remove(tmp);
}

// ── docx::Writer stub ──────────────────────────────────────────────────────

TEST_CASE("Writer stores output path", "[docx]") {
    docweft::docx::Writer writer("/tmp/output.docx");
    REQUIRE(writer.path() == "/tmp/output.docx");
}

TEST_CASE("Writer set_document_xml returns false (stub)", "[docx]") {
    docweft::docx::Writer writer("/tmp/output.docx");
    REQUIRE_FALSE(writer.set_document_xml("<w:document/>"));
}

// ── tmpl::Engine stub ──────────────────────────────────────────────────────

TEST_CASE("Engine render returns template as-is (stub)", "[template]") {
    docweft::tmpl::Engine engine;
    auto result = engine.render("Hello {{ name }}!", "{}");
    // Stub just echoes the template back
    REQUIRE(result == "Hello {{ name }}!");
}

// ── version ────────────────────────────────────────────────────────────────

TEST_CASE("Library version is set", "[core]") {
    REQUIRE(docweft::version == "0.1.0");
}
