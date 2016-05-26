

#pragma once


#include <stdio.h>
#include <tchar.h>


#ifdef _WIN64
#pragma comment(lib, "..\\..\\LibXlsxWriterProj\\Zlib\\x64\\zlib.lib")
#ifdef _DEBUG
#pragma comment(lib, "..\\..\\LibXlsxWriterProj\\x64\\Debug\\LibXlsxWriter.lib")
#else
#pragma comment(lib, "..\\..\\LibXlsxWriterProj\\x64\\Release\\LibXlsxWriter.lib")
#endif
#else
#pragma comment(lib, "..\\..\\LibXlsxWriterProj\\Zlib\\zlib.lib")
#ifdef _DEBUG
#pragma comment(lib, "..\\LibXlsxWriterProj\\Debug\\LibXlsxWriter.lib")
#else
#pragma comment(lib, "..\\..\\LibXlsxWriterProj\\Release\\LibXlsxWriter.lib")
#endif
#endif