#include "vba_internal.h"
#include "vba_exception.h"
#include "vba_Locale.h"

#include "vba_strManipulation.h"

#include "vba_ole_bridge_macros.h"

#define BUNCH_OF_BSTR_CONVERSIONS(declr)																	\
	declr(BSTR, double, __vbaStrR8, VarBstrFromR8, 0, "in '%lf'")					/* double -> BSTR */	\
	declr(BSTR, float, __vbaStrR4, VarBstrFromR4, 0, "in '%f'")						/* float -> BSTR */		\
	declr(BSTR, LONG, __vbaStrI4, VarBstrFromI4, 0, "in '%d'")						/* long -> BSTR */		\
	declr(BSTR, SHORT, __vbaStrI2, VarBstrFromI2, 0, "in '%d'")						/* SHORT -> BSTR */		\
	declr(BSTR, BYTE, __vbaStrUI1, VarBstrFromI2, 0, "in '%d'")						/* BYTE -> BSTR */		\
	declr(BSTR, SHORT, __vbaStrBool, VarBstrFromBool, VAR_LOCALBOOL, "in '%d'")		/* SHORT -> BSTR */		\
	declr(BSTR, CY, __vbaStrCy, VarBstrFromCy , 0, "in '%ld'")						/* CY -> BSTR */		\
	declr(BSTR, DATE, __vbaStrDate, VarBstrFromDate , 0, "in '%ld'")				/* DATE -> BSTR */		\

#pragma warning (disable : 4477) /* This warning is created by converting types in sprintf */

BUNCH_OF_BSTR_CONVERSIONS(DECLARE_VBA_CONVERSION_BRIDGE_TO_OLE_CONVERSION);

/**
 * @brief			Converts a BSTR to an ANSI string.
 * @param			pbstrOut		Pointer to a BSTR (weird, but assume it as a storage).
 *									This function will set the pointer to the ANSI buffer there.
 * @param			bstrSrc			Source BSTR.
 * @return			*pbstrOut always.
 */
EXPORT BSTR __stdcall __vbaStrToAnsi(
	BSTR	* pbstrOut,
	BSTR	bstrSrc
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pbstrOut %.8x, pbstrOut %.8x",
		(unsigned int)pbstrOut,
		(unsigned int)bstrSrc
	);

	if (!pbstrOut)
	{
		vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
		return *pbstrOut;
	}

	int iWideStrSize = 0;

	if (bstrSrc)
	{
		iWideStrSize = strSafeGetLength(bstrSrc);
	}

	/* Get how many bytes we'll need to allocate */
	int size_needed = WideCharToMultiByte(
		CP_ACP,
		0,
		(LPCWCH)bstrSrc,
		iWideStrSize,
		NULL,
		0,
		0,
		0
	);

	DEBUG_WIDE(
		"size_needed = %.8x",
		size_needed
	);

	/* This is really weird, but VB uses a BSTR as storage for an ANSI string */
	*pbstrOut = SysAllocStringByteLen(
		0,
		size_needed
	);

	if (!*pbstrOut)
	{
		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		return *pbstrOut;
	}

	int iChars = WideCharToMultiByte(
		CP_ACP,
		0,
		(LPCWCH)bstrSrc,
		iWideStrSize + 1,
		(LPSTR)*pbstrOut,
		size_needed + 1,
		0,
		0
	);

	DEBUG_WIDE(
		"MultiByteToWideChar %.8x",
		iChars
	);

	return *pbstrOut;
} /* __vbaStrToAnsi */