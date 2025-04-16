cmake_minimum_required(VERSION 3.0)
project(excel_vba)

# Set the C++ standard (C++11 is sufficient, but you may choose higher)
set(CMAKE_CXX_STANDARD 11)

# Locate pybind11 package (install via pip or your package manager)
find_package(pybind11 REQUIRED)

# Create a shared library for the Python module.
add_library(excel_vba MODULE excel_vba.cpp)

# On Windows, the Python extension should have a .pyd extension.
if(WIN32)
    set_target_properties(excel_vba PROPERTIES SUFFIX ".pyd")
endif()

# Link pybind11 to your module.
target_link_libraries(excel_vba PRIVATE pybind11::module)
