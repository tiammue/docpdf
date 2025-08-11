cmake_minimum_required(VERSION 3.16)
project(docpdf VERSION 1.0.0)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Find Qt6 components
find_package(Qt6 REQUIRED COMPONENTS Core Widgets)

# Enable Qt6 features
qt_standard_project_setup()

set(SOURCES
    main.cpp
    mainwindow.cpp
    docpdf.cpp
)

set(HEADERS
    mainwindow.h
    docpdf.h
)

qt_add_executable(docpdf ${SOURCES} ${HEADERS})

target_link_libraries(docpdf Qt6::Core Qt6::Widgets)

# Windows-specific settings
if(WIN32)
    set_target_properties(docpdf PROPERTIES
        WIN32_EXECUTABLE TRUE
    )
endif()