#include "docweft/docx/reader.hpp"

namespace docweft::docx {

Reader::Reader(const std::filesystem::path& docx_path)
    : path_(docx_path)
    , open_(std::filesystem::exists(docx_path)) {}

Reader::~Reader() = default;

bool Reader::is_open() const noexcept { return open_; }

std::string Reader::document_xml() const {
    // Stub — will use libzip + pugixml to extract word/document.xml
    return {};
}

} // namespace docweft::docx
