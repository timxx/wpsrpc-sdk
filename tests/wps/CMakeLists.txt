add_library(test_base STATIC
  testbase.cpp
  testbase.h
)

target_link_libraries(test_base
  Qt5::Test
)

target_include_directories(test_base PRIVATE
  ${SDK_INCS_ROOT}
  ${SDK_INCS_ROOT}/wps
  ${SDK_INCS_ROOT}/common
)

target_precompile_headers(test_base PRIVATE
  "$<$<COMPILE_LANGUAGE:CXX>:${CMAKE_CURRENT_SOURCE_DIR}/pch.h>"
)

function(add_wps_test target)
  add_test_target(${target}
    SRCS ${ARGN}
    LIBS rpcwpsapi_sysqt5 test_base
    INCS ${SDK_INCS_ROOT}/wps
  )
  target_precompile_headers(${target}
    REUSE_FROM test_base
  )
endfunction(add_wps_test)

add_wps_test(comparedocs
  comparedocs.cpp
)
