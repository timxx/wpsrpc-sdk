add_library(test_etbase STATIC
  testetbase.cpp
  testetbase.h
)

target_link_libraries(test_etbase
  Qt5::Test
)

target_include_directories(test_etbase PRIVATE
  ${SDK_INCS_ROOT}
  ${SDK_INCS_ROOT}/et
  ${SDK_INCS_ROOT}/common
)

target_precompile_headers(test_etbase PRIVATE
  "$<$<COMPILE_LANGUAGE:CXX>:${CMAKE_CURRENT_SOURCE_DIR}/pch.h>"
)

function(add_et_test target)
  add_test_target(${target}
    SRCS ${ARGN}
    LIBS rpcetapi_sysqt5 test_etbase
    INCS ${SDK_INCS_ROOT}/et
  )
  target_precompile_headers(${target}
    REUSE_FROM test_etbase
  )
endfunction(add_et_test)

add_et_test(rangevalue
  rangevalue.cpp
)
