// Consumer smoke test for the installed DocWeft CMake package.
// Proves find_package(DocWeft CONFIG) + public headers + linkage all work.
// Does not load or modify any .docx — the goal here is the surface, not behavior.

#include <docweft/error.hpp>
#include <docweft/merger.hpp>

int main() {
    auto merger = docweft::make_docx_merger();
    return merger ? 0 : 1;
}
