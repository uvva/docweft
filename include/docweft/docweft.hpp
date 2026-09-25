#pragma once

#include "docweft/export.h"

#include <filesystem>
#include <string>
#include <string_view>

namespace docweft {

/// Library version
constexpr std::string_view version = "0.1.0";

/// Check whether a file exists at the given path.
[[nodiscard]] DOCWEFT_API bool file_exists(const std::filesystem::path& path);

/// Read entire file contents into a string.
/// Returns empty string if the file cannot be read.
[[nodiscard]] DOCWEFT_API std::string read_file(const std::filesystem::path& path);

} // namespace docweft
