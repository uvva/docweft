#pragma once

#include <string>
#include <string_view>

namespace docweft::tmpl {

/// Placeholder: will render inja/Jinja2-style templates.
class Engine {
public:
    Engine();
    ~Engine();

    /// Render a template string with JSON data (stub).
    [[nodiscard]] std::string render(std::string_view tpl,
                                     std::string_view json_data) const;
};

} // namespace docweft::tmpl
