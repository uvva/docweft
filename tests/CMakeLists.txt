# tests/CMakeLists.txt

include(Catch)

add_executable(docweft_tests
    test_basic.cpp
    test_merger.cpp
    fixture_helpers.cpp
)

target_include_directories(docweft_tests
    PRIVATE
        ${CMAKE_SOURCE_DIR}/src   # for internal headers like docx/merger.hpp
)

# Tests need access to internal symbols (DocxMerger, docx::Reader, etc.).
# When the main library is STATIC, link to it directly.
# When the main library is SHARED, link to the static twin `docweft_internal`.
if(DOCWEFT_BUILD_SHARED)
    set(_tf_test_lib DocWeft::docweft_internal)
else()
    set(_tf_test_lib DocWeft::docweft)
endif()

target_link_libraries(docweft_tests
    PRIVATE
        ${_tf_test_lib}
        Catch2::Catch2WithMain
        libzip::zip               # fixture_helpers creates DOCX archives
        pugixml::pugixml          # internal merger.hpp uses pugi::xml_document
)

# HAVE_* feature flags. In the SHARED build these already propagate through
# docweft_internal (PUBLIC). In the STATIC build docweft hides them
# (PRIVATE), so we repeat them on the test target so conditional test cases
# see the same view of available decoders.
if(DOCWEFT_HAVE_STB)
    target_compile_definitions(docweft_tests PRIVATE DOCWEFT_HAVE_STB=1)
endif()
if(TARGET TIFF::TIFF)
    target_compile_definitions(docweft_tests PRIVATE DOCWEFT_HAVE_TIFF=1)
endif()

# Copy test fixtures next to the binary so tests can find them
set(FIXTURE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/fixtures")
if(EXISTS "${FIXTURE_DIR}")
    add_custom_command(TARGET docweft_tests POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_directory
            "${FIXTURE_DIR}"
            "$<TARGET_FILE_DIR:docweft_tests>/fixtures"
    )
endif()

# Auto-discover Catch2 test cases
catch_discover_tests(docweft_tests
    WORKING_DIRECTORY "$<TARGET_FILE_DIR:docweft_tests>"
)

# End-to-end check: install DocWeft to a staging prefix and build an
# external project against it via find_package(). Covers the install/export
# surface (Targets.cmake, Config.cmake, generated export.h, transitive deps).
add_subdirectory(integration)
