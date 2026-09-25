# examples/CMakeLists.txt
# Two small binaries that show how to consume the DocWeft library.

# ── generate_template — uses libzip directly (no docweft dependency) ────
add_executable(generate_template generate_template.cpp)
target_link_libraries(generate_template PRIVATE libzip::zip)

# ── basic_report — uses the docweft public API only ────────────────────
add_executable(basic_report basic_report.cpp)
target_link_libraries(basic_report PRIVATE DocWeft::docweft)

set(_tf_example_targets generate_template basic_report)

# ── native_pdf_report — demonstrates the native PoDoFo PDF backend, i.e.
# save(".pdf") producing real output with no Microsoft Word or LibreOffice
# installed at all. Only built when that backend is compiled in — see
# PLAN.md — since the whole point of the example is exercising that path;
# without it, save() would just throw NoConverter once the demo disables
# LibreOffice detection.
if(DOCWEFT_HAVE_PODOFO)
    add_executable(native_pdf_report native_pdf_report.cpp)
    target_link_libraries(native_pdf_report PRIVATE DocWeft::docweft libzip::zip)
    list(APPEND _tf_example_targets native_pdf_report)
endif()

# Place example binaries + the generated template together so running
# `./basic_report` out-of-the-box just works.
set_target_properties(${_tf_example_targets} PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}"
)

# ── Windows: copy docweft.dll (and its transitive runtime DLLs) ────────
# next to every example executable so they run out-of-the-box without
# touching PATH. No-op on Linux/macOS: rpath handles shared-lib lookup there.
if(WIN32 AND DOCWEFT_BUILD_SHARED)
    set(_tf_dll_copy_targets basic_report)
    if(TARGET native_pdf_report)
        list(APPEND _tf_dll_copy_targets native_pdf_report)
    endif()
    foreach(_tf_example ${_tf_dll_copy_targets})
        add_custom_command(TARGET ${_tf_example} POST_BUILD
            COMMAND ${CMAKE_COMMAND} -E copy_if_different
                    $<TARGET_RUNTIME_DLLS:${_tf_example}>
                    $<TARGET_FILE_DIR:${_tf_example}>
            COMMAND_EXPAND_LISTS
            COMMENT "Copying runtime DLLs next to ${_tf_example}"
        )
    endforeach()
endif()

# Generate the sample template.docx at build time so users can immediately
# run ./basic_report template.docx report.docx without any extra steps.
add_custom_command(
    OUTPUT  "${CMAKE_CURRENT_BINARY_DIR}/template.docx"
    COMMAND generate_template "${CMAKE_CURRENT_BINARY_DIR}/template.docx"
    DEPENDS generate_template
    WORKING_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}"
    COMMENT "Generating examples/template.docx"
    VERBATIM
)

# Copy the committed sample logo.png next to the binaries so basic_report
# can embed it out-of-the-box. The file is tracked in git (75 bytes) — see
# examples/logo.png. POST_BUILD-on-basic_report ensures it lands whenever
# someone rebuilds the binary.
add_custom_command(
    OUTPUT  "${CMAKE_CURRENT_BINARY_DIR}/logo.png"
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "${CMAKE_CURRENT_SOURCE_DIR}/logo.png"
            "${CMAKE_CURRENT_BINARY_DIR}/logo.png"
    DEPENDS "${CMAKE_CURRENT_SOURCE_DIR}/logo.png"
    COMMENT "Copying examples/logo.png to build dir"
    VERBATIM
)

add_custom_target(example_template ALL
    DEPENDS "${CMAKE_CURRENT_BINARY_DIR}/template.docx"
            "${CMAKE_CURRENT_BINARY_DIR}/logo.png"
)
