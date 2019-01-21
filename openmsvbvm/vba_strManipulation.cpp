#include "vba_internal.h"
#include "vba_exception.h"

#include <cstring>
#include <string>

#include "vba_objManipulation.h"
#include "vba_varManipulation.h"

/**
 * @brief			Gets the string length of a BSTR.
 * @param			bstrIn				Input string.
 * @returns			Size of the string if the BSTR is valid, 0 otherwise.
 */
unsigned int __stdcall strSafeGetLength(
	BSTR			bstrIn
)
{
	if (bstrIn)
	{
		// The string length is stored in the previous 4 bytes
		// of the starting character of the BSTR. It's weird.
		uint32_t		*puintSize = (uint32_t *)bstrIn - 1;

		// Check if we're going to create an exception if we read the size.
		if (IsBadReadPtr(puintSize, sizeof(uint32_t)))
		{
			return 0;
		}

		// Size is in bytes, not in wide-chars, so divide by two.
		return *puintSize / 2;
	}

	return 0;
} /* strSafeGetLength */

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcReplace(
	BSTR			bstrExpression,
	BSTR			bstrFind,
	BSTR			bstrReplace,
	int				iStart,
	int				iCount,
	int				iCompareMethod
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"bstrExpression %.8x '%ls', bstrFind %.8x '%ls', bstrReplace %.8x '%ls', iStart %.8x, iCount %.8x, iCompareMethod %.8x",
		(unsigned int)bstrExpression,
		(wchar_t*)bstrExpression,
		(unsigned int)bstrFind,
		(wchar_t*)bstrFind,
		(unsigned int)bstrReplace,
		(wchar_t*)bstrReplace,
		iStart,
		iCount,
		iCompareMethod
	);
	return bstrExpression;
} /* rtcReplace */

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			BSTR String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcUpperCaseBstr(
	BSTR			strIn
)
{
	UINT uStrSize = strSafeGetLength(strIn);
	BSTR bstrRet;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"strIn %.8x, size %.8x",
		(unsigned int)strIn,
		uStrSize
	);

	bstrRet = SysAllocStringLen(
		strIn,
		uStrSize
	);

	if (!bstrRet)
	{
		DEBUG_WIDE(
			"SysAllocStringLen failed, err = %.8x",
			GetLastError()
		);

		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		return nullptr;
	}

	CharUpperBuffW(bstrRet, uStrSize + 1);

	DEBUG_WIDE("ret = '%ls'", bstrRet);

	return bstrRet;
} /* rtcUpperCaseBstr */

/**
 * @brief			Copies a BSTR and converts their characters to lowercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcLowerCaseBstr(
	BSTR			strIn
)
{
	UINT uStrSize = strSafeGetLength(strIn);
	BSTR bstrRet;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"strIn %.8x, size %.8x",
		(unsigned int)strIn,
		uStrSize);

	bstrRet = SysAllocStringLen(
		strIn,
		uStrSize
	);

	if (!bstrRet)
	{
		DEBUG_WIDE(
			"SysAllocStringLen failed, err = %.8x",
			(unsigned int)GetLastError()
		);

		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		return NULL;
	}

	CharLowerBuffW(bstrRet, uStrSize + 1);

	DEBUG_WIDE("ret = '%ls'", bstrRet);

	return bstrRet;
} /* rtcLowerCaseBstr */

/**
 * @brief			Concatenates two BSTRs.
 * @param			bstrRight		Right operand for the concatenation.
 * @param			bstrLeft		Left operand for the concatenation.
 * @returns			Concatenated BSTR on success, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrCat(
	BSTR			bstrLeft,
	BSTR			bstrRight
)
{
	HRESULT result;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"bstrRight %.8x, bstrLeft %.8x",
		(unsigned int)bstrRight,
		(unsigned int)bstrLeft
	);

	if (bstrRight && bstrLeft)
	{
		DEBUG_WIDE(
			"bstrRight '%ls', bstrLeft '%ls'",
			(wchar_t*)bstrRight,
			(wchar_t*)bstrLeft
		);
	}

	BSTR	ret;

	result = VarBstrCat(
		bstrRight,
		bstrLeft,
		&ret
	);

	if (result < 0)
	{
		DEBUG_WIDE("result %.8x", result);
		vbaRaiseException(vbaErrorFromHRESULT(result));
		return NULL;
	}

	DEBUG_WIDE(
		"result %.8x, '%ls'",
		result,
		(wchar_t*)&ret
	);

	return ret;
} /* __vbaStrCat */

/**
 * @brief			Moves one BSTR to the pointer of another (previously freed if needed) BSTR.
 * @param			pbstrDest		Where the source BSTR value will be copied to. Freed if not zero.
 * @param			bstrSrc			Source BSTR.
 * @returns			Same value as pbstrSrc.
 */
EXPORT BSTR __fastcall __vbaStrMove(
	BSTR			* pbstrDest,
	BSTR			bstrSrc
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pbstrDest %.8x, bstrSrc %.8x",
		(unsigned int)pbstrDest,
		(unsigned int)bstrSrc
	);

	if (*pbstrDest)
	{
		DEBUG_WIDE(
			"pbstrDest contained '%ls'",
			(wchar_t*)pbstrDest
		);

		SysFreeString(*pbstrDest);
	}

	if (bstrSrc)
	{
		DEBUG_WIDE(
			"bstrSrc '%ls'",
			(wchar_t*)bstrSrc
		);
	}

	*pbstrDest = bstrSrc;
	return bstrSrc;
} /* __vbaStrMove */

/**
 * @brief			Converts a Variant (including VT_ERROR) to a BSTR using OLE's VariantChangeType.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrErrVarCopy(
	VARIANTARG		*pvargIn
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	VARIANT			vargRet;
	VariantInit(&vargRet);

	if (!pvargIn)
	{
		return nullptr;
	}

	DEBUG_WIDE(
		"pvargIn %.8x, pvargIn->vt %.8x",
		(unsigned int)pvargIn,
		(unsigned int)pvargIn->vt
	);

	HRESULT hr = VariantChangeType(
		&vargRet,
		pvargIn,
		VARIANT_ALPHABOOL | VARIANT_LOCALBOOL,
		VT_BSTR
	);

	if (hr != S_OK)
	{
		return nullptr;
	}

	UINT		uStrSize = strSafeGetLength(vargRet.bstrVal);

	DEBUG_WIDE(
		"vargRet.bstrVal %.8x, size %.8x",
		(unsigned int)vargRet.bstrVal,
		uStrSize
	);

	BSTR bstrRet = SysAllocStringLen(
		vargRet.bstrVal,
		uStrSize
	);

	__vbaFreeVar(&vargRet);

	if (!bstrRet)
	{
		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
	}

	return bstrRet;
} /* __vbaStrErrVarCopy */

/**
 * @brief			Copies one BSTR to the pointer of another (previously freed if needed) BSTR.
 * @param			strOut			Where the source BSTR value will be copied to. Freed if not zero.
 * @param			strIn			Source BSTR.
 * @returns			Contents of strOut.
 */
EXPORT BSTR __fastcall __vbaStrCopy(
	BSTR			* strOut,
	BSTR			strIn
)
{
	BSTR bstrRet = nullptr;
	UINT uStrSize = 0;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (strIn)
	{
		uStrSize = strSafeGetLength(strIn);

		bstrRet = SysAllocStringLen(
			strIn,
			uStrSize
		);

		if (!bstrRet)
		{
			vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		}
	}

	DEBUG_WIDE(
		"strIn %.8x, strOut %.8x, *strOut %.8x, size %.8x, ret %.8x",
		(unsigned int)strIn,
		(unsigned int)strOut,
		(unsigned int)*strOut,
		uStrSize,
		(unsigned int)bstrRet
	);

	if (*strOut)
	{
		SysFreeString(*strOut);
	}

	*strOut = bstrRet;

	DEBUG_WIDE(
		"ret = %.8x",
		(unsigned int)bstrRet
	);

	return bstrRet;
} /* __vbaStrCopy */

/**
 * @brief			Converts a Variant (non including errors) to a BSTR.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, NULL on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrVarCopy(
	VARIANTARG		* pvargIn
)
{
	return __vbaStrErrVarCopy(pvargIn);
} /* __vbaStrVarCopy */

/**
 * @brief			Frees a BSTR via its pointer, and nulls it.
 * @param			pbstrIn			Pointer to a BSTR to free and null it.
 */
EXPORT void __fastcall __vbaFreeStr(
	BSTR			*pbstrIn
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (pbstrIn)
	{
		DEBUG_WIDE(
			"*pbstrIn %.8x",
			(unsigned int)*pbstrIn
		);

		if (*pbstrIn)
		{
			DEBUG_WIDE(
				"*pbstrIn '%ls'",
				(wchar_t*)*pbstrIn
			);

			SysFreeString(*pbstrIn);
			*pbstrIn = 0;
		}
	}
} /* __vbaFreeStr */

/**
 * @brief			Frees a list of BSTRs via their pointers, and nulls them.
 * @param			argCount		Count of elements.
 * @param			...				Pointers to BSTRs to free and null them.
 */
EXPORT void __cdecl __vbaFreeStrList(
	unsigned int	argCount,
	...
)
{
	BSTR		*pbstrElement;

	va_list		args;
	va_start(args, argCount);

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	while (argCount--)
	{
		pbstrElement = va_arg(args, BSTR*);
		DEBUG_WIDE(
			"arg() %.8x, argsRemaining %.8x",
			(unsigned int)pbstrElement,
			argCount
		);

		__vbaFreeStr(pbstrElement);
	}

	va_end(args);
} /* __vbaFreeStrList */