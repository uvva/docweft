#pragma once

#include <filesystem>
#include <string_view>

namespace docweft::docx {

/// Placeholder: will create/modify a .docx archive.
class Writer {
public:
    explicit Writer(const std::filesystem::path& output_path);
    ~Writer();

    /// Returns the target output path.
    [[nodiscard]] const std::filesystem::path& path() const noexcept;

    /// Stub: write document.xml content into the archive.
    bool set_document_xml(std::string_view xml);

private:
    std::filesystem::path path_;
};

} // namespace docweft::docx
