#pragma once

#include <windows.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <stdbool.h>
#include <stdlib.h>

#include <oaidl.h>


#define DEBUG_STORAGE_SIZE							300
#define DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED()		char	debugA[DEBUG_STORAGE_SIZE];
#define DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED()		wchar_t debugW[DEBUG_STORAGE_SIZE];

#define WIDE2(x) L##x
#define WIDECHAR(x) WIDE2(x)

#define WIDE_FUNCTION WIDECHAR(__FUNCTION__)

#define DEBUG_ASCII(fmt, ...)						snprintf(debugA, DEBUG_STORAGE_SIZE - 1, "%s: " fmt, __FUNCTION__, __VA_ARGS__); OutputDebugStringA(debugA);
#define DEBUG_ASCII_OBJ(fmt, ...)					snprintf(debugA, DEBUG_STORAGE_SIZE - 1, "%s(%.8x): " fmt, __FUNCTION__, (unsigned long)this, __VA_ARGS__); OutputDebugStringA(debugA);
#define DEBUG_WIDE(fmt, ...)						swprintf(debugW, DEBUG_STORAGE_SIZE - 1, L"%s: " L##fmt, WIDE_FUNCTION, __VA_ARGS__); OutputDebugStringW(debugW);
#define DEBUG_WIDE_OBJ(fmt, ...)					swprintf(debugW, DEBUG_STORAGE_SIZE - 1, L"%s(%.8x): " L##fmt, WIDE_FUNCTION, (unsigned long)this, __VA_ARGS__); OutputDebugStringW(debugW);
#define DEBUG_WIDE_GUID(guid, prepend)				{ OLECHAR* guidString; StringFromCLSID(guid, &guidString); DEBUG_WIDE(prepend ": %ls", guidString); CoTaskMemFree(guidString); }




#define ARGS(...)		__VA_ARGS__


#ifdef __cplusplus
#define CEXTERN extern "C"
#else
#define CEXTERN 
#endif


#ifdef __cplusplus
#define EXPORT extern "C" __declspec (dllexport)
#else
#define EXPORT __declspec (dllexport)
#endif