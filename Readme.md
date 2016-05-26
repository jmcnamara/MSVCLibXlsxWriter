# MSVCLibXlsxWriter


MSVCLibXlsxWriter is a MSVC project to build a Windows DLL for [libxlsxwriter][libxlsxwriter] a C library for creating Excel XLSX files.

[libxlsxwriter]: https://github.com/jmcnamara/libxlsxwriter

![demo image](http://libxlsxwriter.github.io/demo.png)

Libxlsxwriter is a C library that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.

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

See the [full libxlsxwriter documentation](http://libxlsxwriter.github.io) for a getting started guide, a tutorial, the main API documentation and examples. Or see the [source code on GitHub](https://github.com/jmcnamara/libxlsxwriter).


## Building a Windows DLL of libxlsxwriter

The MSVCLibXlsxWriter repository contains 3 directories:

- **LibXlsxWriterProj**: A MSVC project to build a `LibXlsxWriter.dll` from the libxlsxwriter source code. The directory also contains a pre-built `Zlib.dll` file.
- **libxlsxwriter**: The libxlsxwriter source code in a git submodule.
- **ExampleExe**: A libxlsxwriter sample application built as a console application that requires the `LibXlsxWriter.dll` and `Zlib.dll` files.

The libxlsxwriter directory is a Git submodule. It isn't included when you do a Git clone of MSVCLibXlsxWriter. In order to get the submodule you can clone the project recursively as follows:

    git clone --recursive https://github.com/jmcnamara/MSVCLibXlsxWriter.git

Or:

    git clone https://github.com/jmcnamara/MSVCLibXlsxWriter.git
    cd MSVCLibXlsxWriter/
    git submodule init
    git submodule update

Open the `LibXlsxWriterProj/LibXlsxWriter.sln` project in MS Visual Studio and build the DLL using the "Build -> Build Solution" menu item.

In the default configuration this will build an x64 debug LibXlsxWriter `.lib` and `.dll` in:

    MSVCLibXlsxWriter\LibXlsxWriterProj\x64\Debug


## Building a console application using the LibXlsxWriter.lib

Ensure that `LibXlsxWriter.lib` was built correctly in the previous steps.

Open the `ExampleExe/ExampleExe.sln` project in MS Visual Studio and build the DLL using the "Build -> Build Solution" menu item.

In the default configuration this will build an x64 exe file as follows:

    MSVCLibXlsxWriter\ExampleExe\x64\Debug\ExampleExe.exe

To run the application copy the `LibXlsxWriter.dll` and `Zlib.dll` file from the `MSVCLibXlsxWriter\LibXlsxWriterProj` sub-directories to the same directory as the executable. You can then run the application by double clicking on it in File Explorer or by opening a CMD console and running it from the directory.

Once the program has run it will create a `chart_column.xlsx` file based on the default sample application in ExampleExe.cpp. You can run other libxlsxwriter example programs by copying the code from one of the `libxlsxwriter\example\*.c` programs.
