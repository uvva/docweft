# tests/integration/CMakeLists.txt
#
# End-to-end check of the installed CMake package. Three ctest fixtures run
# in sequence:
#
#   integration_install            cmake --install → staging prefix
#     └─ integration_consumer_configure   cmake -S <consumer> -B <build>
#           └─ integration_consumer_build   cmake --build <build>
#
# Failure at any step means the exported package (DocWeftConfig.cmake,
# DocWeftTargets.cmake, generated export.h, find_dependency chain) is
# broken for downstream users.

set(_install_stage "${CMAKE_CURRENT_BINARY_DIR}/install_stage")
set(_consumer_src  "${CMAKE_CURRENT_SOURCE_DIR}/consumer_find_package")
set(_consumer_bld  "${CMAKE_CURRENT_BINARY_DIR}/consumer-build")

# Forward vcpkg toolchain / triplet so the consumer sees the same dependency
# view. Without this the consumer's find_dependency(pugixml) etc. would miss
# on a vcpkg-only Windows host.
set(_forward_args "")
if(CMAKE_TOOLCHAIN_FILE)
    list(APPEND _forward_args "-DCMAKE_TOOLCHAIN_FILE=${CMAKE_TOOLCHAIN_FILE}")
endif()
if(VCPKG_TARGET_TRIPLET)
    list(APPEND _forward_args "-DVCPKG_TARGET_TRIPLET=${VCPKG_TARGET_TRIPLET}")
endif()

add_test(NAME integration_install
    COMMAND ${CMAKE_COMMAND}
        --install ${CMAKE_BINARY_DIR}
        --prefix  ${_install_stage}
        --config  $<CONFIG>
)
set_tests_properties(integration_install PROPERTIES
    FIXTURES_SETUP DocWeftInstalled
    LABELS         "integration"
)

add_test(NAME integration_consumer_configure
    COMMAND ${CMAKE_COMMAND}
        -S ${_consumer_src}
        -B ${_consumer_bld}
        -G ${CMAKE_GENERATOR}
        -DCMAKE_PREFIX_PATH=${_install_stage}
        -DCMAKE_BUILD_TYPE=$<CONFIG>
        ${_forward_args}
)
set_tests_properties(integration_consumer_configure PROPERTIES
    FIXTURES_REQUIRED DocWeftInstalled
    FIXTURES_SETUP    DocWeftConsumerConfigured
    LABELS            "integration"
)

add_test(NAME integration_consumer_build
    COMMAND ${CMAKE_COMMAND}
        --build  ${_consumer_bld}
        --config $<CONFIG>
)
set_tests_properties(integration_consumer_build PROPERTIES
    FIXTURES_REQUIRED DocWeftConsumerConfigured
    LABELS            "integration"
)
