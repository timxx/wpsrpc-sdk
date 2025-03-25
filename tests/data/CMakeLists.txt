set(data_files
    chart.xlsx
)

set(_outdir "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}/data")

foreach(_file ${data_files})
    set(_out_file "${_outdir}/${_file}")

    add_custom_command(
        OUTPUT ${_out_file}
        COMMAND ${CMAKE_COMMAND} -E copy ${CMAKE_CURRENT_SOURCE_DIR}/${_file} ${_out_file}
        DEPENDS ${CMAKE_CURRENT_SOURCE_DIR}/${_file}
        COMMENT "Copying ${_file} to data directory"
    )
    list(APPEND _data_output_files ${_out_file})
endforeach()

add_custom_target(test_data ALL
    DEPENDS ${_data_output_files}
)
