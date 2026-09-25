#pragma once

#include <filesystem>
#include <string>

namespace docweft::docx {

/// Placeholder: will read and parse a .docx archive.
class Reader {
public:
    explicit Reader(const std::filesystem::path& docx_path);
    ~Reader();

    /// Returns true if the archive was opened successfully.
    [[nodiscard]] bool is_open() const noexcept;

    /// Returns raw XML content of word/document.xml (stub).
    [[nodiscard]] std::string document_xml() const;

private:
    std::filesystem::path path_;
    bool open_ = false;
};

} // namespace docweft::docx
