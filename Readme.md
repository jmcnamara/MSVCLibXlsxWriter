# MSVCLibXlsxWriter


MSVCLibXlsxWriter is a MSVC project to build a Windows DLL for
[libxlsxwriter][lxw_git] a C library for creating Excel XLSX files.

[lxw_git]: https://github.com/jmcnamara/libxlsxwriter

![demo image](http://libxlsxwriter.github.io/demo.png)

Libxlsxwriter is a C library that can be used to write text, numbers, formulas
and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.

It supports features such as:

- 100% compatible Excel XLSX files.
- Full Excel formatting.
- Merged cells.
- Defined names.
- Autofilters.
- Charts.
- Worksheet PNG/JPEG images.
- Memory optimization mode for writing large files.
- Source code available on [GitHub](https://github.com/jmcnamara/libxlsxwriter).
- FreeBSD ref license.
- ANSI C.
- Works with GCC 4.x, GCC 5.x, Clang, Xcode, MSVC 2015, ICC and TCC.
- Works on Linux, FreeBSD, OS X, iOS and Windows.
- The only dependency is on `zlib`.

See the full [libxlsxwriter documentation][lxw_docs] for a getting started
guide, a tutorial, the main API documentation and examples. Or browse the
[source code on GitHub][lxw_git].

[lxw_docs]: http://libxlsxwriter.github.io


## Building a Windows DLL of libxlsxwriter

The MSVCLibXlsxWriter repository contains 3 directories:

- **LibXlsxWriterProj**: A MSVC project to build a `LibXlsxWriter.dll` from
    the libxlsxwriter source code. The directory also contains a pre-built
    `Zlib.dll` file.

- **ExampleExe**: A libxlsxwriter sample application built as a console
    application that requires the `LibXlsxWriter.dll` and `Zlib.dll` files.

- **libxlsxwriter**: The libxlsxwriter source code in a git submodule, see
    below.

The `libxlsxwriter` directory is a Git submodule. This means that it isn't
included when you do a standard Git clone of MSVCLibXlsxWriter. In order to
get the submodule as well as the project code you must clone the project
recursively as follows:

    git clone --recursive https://github.com/jmcnamara/MSVCLibXlsxWriter.git

Or update it explicitly as follows:

    git clone https://github.com/jmcnamara/MSVCLibXlsxWriter.git
    cd MSVCLibXlsxWriter/
    git submodule init
    git submodule update

This version of MSVCLibXlsxWriter contains libxlsxwriter version 0.3.4.



To build the DLL of the library open the `LibXlsxWriterProj/LibXlsxWriter.sln`
project in MS Visual Studio and build the solution using the "Build -> Build
Solution" menu item.

In the default configuration this will build an x64 debug LibXlsxWriter `.lib`
and `.dll` in:

    MSVCLibXlsxWriter\LibXlsxWriterProj\x64\Debug


## Building a console application using the LibXlsxWriter.lib

Ensure that `LibXlsxWriter.lib` was built correctly in the previous steps.

To build the example executable open the `ExampleExe/ExampleExe.sln` project
in MS Visual Studio and build the solution using the "Build -> Build Solution"
menu item.

In the default configuration this will build the following x64 exe file:

    MSVCLibXlsxWriter\ExampleExe\x64\Debug\ExampleExe.exe

To run the application copy the `LibXlsxWriter.dll` and `Zlib.dll` files from
the `MSVCLibXlsxWriter\LibXlsxWriterProj` sub-directories to the same
directory as the executable. You can then run the application by double
clicking on it in File Explorer or by opening a CMD console and running it
from the directory.

Once the program has run it will create a `chart_column.xlsx` file based on
the default sample application in ExampleExe.cpp. You can run other
libxlsxwriter example programs by copying the code from one of the
`libxlsxwriter\example\*.c` programs.
