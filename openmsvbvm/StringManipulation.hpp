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
 * @brief			Returns pvarSrc's value as a BSTR, avoiding a copy when it's
 *					already a plain VT_BSTR (confirmed live via raw-byte disassembly
 *					of the real msvbvm60.dll, ordinal 438 -- the first stack parameter
 *					[esp+4] is the one zeroed out on the VT_BSTR fast path, i.e.
 *					pOutOwned comes FIRST and pvarSrc SECOND, the opposite of this
 *					project's own earlier, incorrect guess at the order): if
 *					pvarSrc->vt == VT_BSTR, returns pvarSrc->bstrVal directly and sets
 *					*pOutOwned to nullptr (the returned pointer is BORROWED from
 *					pvarSrc -- the caller must NOT free it). Otherwise coerces via
 *					__vbaStrVarCopy and sets *pOutOwned to that same freshly-allocated
 *					result (the caller now OWNS it and is responsible for freeing it).
 *					Either way, the function's own return value is the BSTR to
 *					actually use.
 * @param			pOutOwned		Set to nullptr (borrowed) or the owned BSTR (must
 *									be freed by the caller) -- see above.
 * @param			pvarSrc			Value to read as a string.
 * @returns			The BSTR value (borrowed or owned, see pOutOwned), or NULL if
 *					pvarSrc is null or the coercion fails.
 */
EXPORT BSTR __stdcall __vbaStrVarVal(
	BSTR			*pOutOwned,
	VARIANTARG		*pvarSrc
);

/**
 * @brief			Gets the length (in characters) of a BSTR. Confirmed via real
 *					msvbvm60.dll disassembly (ordinal 327): reads the BSTR's own
 *					length prefix (same computation as this project's own
 *					strSafeGetLength) and returns 0 for a null BSTR.
 * @param			bstrIn			Input string.
 * @returns			Length in characters, or 0 if bstrIn is null.
 */
EXPORT int __stdcall __vbaLenBstr(
	BSTR			bstrIn
);

/**
 * @brief			Takes ownership of a Variant's string value, coercing it to
 *					VT_BSTR in place first if it isn't one already, and empties the
 *					Variant to VT_EMPTY WITHOUT freeing the string -- ownership of
 *					the returned BSTR transfers to the caller. Confirmed via real
 *					msvbvm60.dll disassembly (ordinal 437).
 * @param			pvarg			Variant to take the string value from. Left as
 *									VT_EMPTY on return.
 * @returns			The BSTR value (now owned by the caller), or NULL on failure.
 */
EXPORT BSTR __stdcall __vbaStrVarMove(
	VARIANTARG		*pvarg
);

/**
 * @brief			Implements Asc$(): the character code of a string's first
 *					character, using the current ANSI code page. Confirmed via real
 *					msvbvm60.dll disassembly (ordinal 516) for the common
 *					single-byte-codepage path; raises an exception for a null or
 *					empty string. Does not yet replicate the real DLL's separate
 *					handling of DBCS (East Asian code page) lead/trail byte pairs.
 * @param			bstrIn			Input string.
 * @returns			ANSI character code of the first character.
 */
EXPORT short __stdcall rtcAnsiValueBstr(
	BSTR			bstrIn
);

/**
 * @brief			Implements Mid$(string, start, [length]) into a Variant result
 *					slot. Confirmed via real msvbvm60.dll disassembly (ordinal 632):
 *					pvargString must be VT_BSTR or VT_BSTR|VT_BYREF (else Type
 *					Mismatch, same unwrap contract as UCase$/LCase$/Trim$'s shared
 *					internal helper); iStart arrives as a plain int (Start is
 *					required, so the VB6 compiler coerces it before the call); Length
 *					stays boxed as a Variant specifically so a VT_ERROR Variant can
 *					mean "omitted" (Length is this function's only optional
 *					argument). This reimplements that external contract in terms of
 *					this project's own Variant/BSTR helpers rather than porting the
 *					real DLL's internal helper chain (rtcMidCharBstr/rtcMidBstr)
 *					instruction-for-instruction.
 * @param			pvargResult		Destination Variant, receives the VT_BSTR result.
 * @param			pvargString		Source string Variant (BSTR or BYREF BSTR).
 * @param			iStart			1-based start position.
 * @param			pvargLength		Length Variant, or a VT_ERROR Variant if Length
 *									was omitted (meaning "to the end of the string").
 * @returns			pvargResult.
 */
EXPORT VARIANTARG * __stdcall rtcMidCharVar(
	VARIANTARG		*pvargResult,
	VARIANTARG		*pvargString,
	int				iStart,
	VARIANTARG		*pvargLength
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