#include "docweft/docx/writer.hpp"

namespace docweft::docx {

Writer::Writer(const std::filesystem::path& output_path)
    : path_(output_path) {}

Writer::~Writer() = default;

const std::filesystem::path& Writer::path() const noexcept { return path_; }

bool Writer::set_document_xml([[maybe_unused]] std::string_view xml) {
    // Stub — will use libzip to write xml into the archive
    return false;
}

} // namespace docweft::docx
