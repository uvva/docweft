#pragma once

#include "textfabric/export.h"

#include <stdexcept>
#include <string>
#include <string_view>

namespace textfabric {

/// Error codes returned across the library boundary.
/// Values are stable; new codes append at the end.
enum class ReportError : int {
    None                  = 0,
    CantOpenTemplate      = 1,   ///< file cannot be opened or is not a DOCX
    CantCopyDocxTemplate  = 2,   ///< template is corrupted / not valid OOXML
    InvalidBookmark       = 3,   ///< requested bookmark not found
    InvalidField          = 4,   ///< requested field inside bookmark not found
    SaveFailed            = 5,   ///< output could not be written
    NoConverter           = 6,   ///< no conversion tool available for .pdf/.html
    NotImplemented        = 7,   ///< feature scheduled for later stage
    Unknown               = 99,
};

[[nodiscard]] TEXTFABRIC_API std::string_view to_string(ReportError e) noexcept;

/// Exception thrown across the library boundary.
/// External code should catch this and translate to its own error type.
class TEXTFABRIC_API ReportException : public std::runtime_error {
public:
    ReportException(ReportError code, std::string msg)
        : std::runtime_error(std::move(msg))
        , code_(code) {}

    [[nodiscard]] ReportError code() const noexcept { return code_; }

private:
    ReportError code_;
};

} // namespace textfabric
