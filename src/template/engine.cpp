#include "docweft/template/engine.hpp"

namespace docweft::tmpl {

Engine::Engine()  = default;
Engine::~Engine() = default;

std::string Engine::render([[maybe_unused]] std::string_view tpl,
                           [[maybe_unused]] std::string_view json_data) const {
    // Stub — will use inja + nlohmann_json
    return std::string(tpl);
}

} // namespace docweft::tmpl
