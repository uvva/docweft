#include "docweft/error.hpp"

namespace docweft {

std::string_view to_string(ReportError e) noexcept {
    switch (e) {
        case ReportError::None:                 return "None";
        case ReportError::CantOpenTemplate:     return "CantOpenTemplate";
        case ReportError::CantCopyDocxTemplate: return "CantCopyDocxTemplate";
        case ReportError::InvalidBookmark:      return "InvalidBookmark";
        case ReportError::InvalidField:         return "InvalidField";
        case ReportError::SaveFailed:           return "SaveFailed";
        case ReportError::NoConverter:          return "NoConverter";
        case ReportError::NotImplemented:       return "NotImplemented";
        case ReportError::Unknown:               break;
    }
    return "Unknown";
}

} // namespace docweft
