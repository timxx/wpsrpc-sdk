cmake_minimum_required(VERSION 3.0.2 FATAL_ERROR)

project(wpsrpc-sdk)

find_package(Qt5Core REQUIRED)
find_package(Qt5Test REQUIRED)

include(CTest)

set(SDK_INCS_ROOT "${CMAKE_SOURCE_DIR}/include")
set(SDK_LIBS_ROOT "${CMAKE_SOURCE_DIR}/lib/${CMAKE_SYSTEM_PROCESSOR}")

set(CMAKE_AUTOMOC ON)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}/bin")
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}")
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}")

set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-attributes")

macro(ADD_TEST_TARGET target)
  cmake_parse_arguments(TEST_ARG
    ""
    ""
    "SRCS;LIBS;INCS;"
    ${ARGN}
  )
  add_executable(${target}
    ${TEST_ARG_SRCS}
  )
  target_link_libraries(${target}
    Qt5::Test
    ${TEST_ARG_LIBS}
  )

  target_include_directories(${target} PRIVATE
    ${TEST_ARG_INCS}
    ${SDK_INCS_ROOT}
    ${SDK_INCS_ROOT}/common
  )
  target_link_directories(${target} PRIVATE
    ${SDK_LIBS_ROOT}
  )
  add_test(NAME ${target} COMMAND ${target})
endmacro()

add_subdirectory(tests)
