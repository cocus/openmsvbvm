#pragma once
#include "vba_internal.h"

/**
 * @brief			Gets the BSTR string length.
 * @param			bstrIn				Input string.
 * @returns			Size of the string if the BSTR is valid, 0 otherwise.
 */
unsigned int __stdcall strSafeGetLength(
	BSTR			bstrIn
);

/**
 * @brief			Converts a Variant (including errors) to a BSTR.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrErrVarCopy(
	VARIANTARG		*pvargIn
);

/**
 * @brief			Converts a Variant (non including errors) to a BSTR.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrVarCopy(
	VARIANTARG		*pvargIn
);

/**
 * @brief			Frees a list of BSTRs via their pointers, and nulls them.
 * @param			argCount		Count of elements.
 * @param			...				Pointers to BSTRs to free and null them.
 */
EXPORT void __cdecl __vbaFreeStrList(
	unsigned int argCount,
	...
);

/**
 * @brief			Frees a BSTR via its pointer, and nulls it.
 * @param			pbstrIn			Pointer to a BSTR to free and null it.
 */
EXPORT void __fastcall __vbaFreeStr(
	BSTR		*pbstrIn
);

/**
 * @brief			Moves one BSTR to the pointer of another (previously freed if needed) BSTR.
 * @param			pbstrDest		Where the source BSTR value will be copied to. Freed if not zero.
 * @param			bstrSrc			Source BSTR.
 * @returns			Same value as pbstrSrc.
 */
EXPORT BSTR __fastcall __vbaStrMove(
	BSTR		*pbstrDest,
	BSTR		bstrSrc
);

/**
 * @brief			Concatenates two BSTRs.
 * @param			bstrRight		Right operand for the concatenation.
 * @param			bstrLeft		Left operand for the concatenation.
 * @returns			Same value as bstrLeft on success, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrCat(
	BSTR bstrLeft,
	BSTR bstrRight
);

/**
 * @brief			Copies a BSTR and converts their characters to lowercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcLowerCaseBstr(
	OLECHAR *strIn
);

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcUpperCaseBstr(
	BSTR strIn
);

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcReplace(
	BSTR		bstrExpression,
	BSTR		bstrFind,
	BSTR		bstrReplace,
	int			iStart,
	int			iCount,
	int			iCompareMethod
);