#pragma once

#include <windows.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <stdbool.h>
#include <stdlib.h>

#include <oaidl.h>

#define ARGS(...) __VA_ARGS__

#ifdef __cplusplus
#define CEXTERN extern "C"
#else
#define CEXTERN
#endif

#ifdef __cplusplus
#define EXPORT extern "C" __declspec(dllexport)
#else
#define EXPORT __declspec(dllexport)
#endif
