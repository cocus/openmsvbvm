#include "vba_internal.h"
#include "vba_exception.h"
#include "vba_Locale.h"

#include "vba_varManipulation.h"
#include "vba_strManipulation.h"
#include "vba_objManipulation.h"

#include "vba_ole_bridge_macros.h"

#define BUNCH_OF_VAR_MANIPULATIONS(declr)	\
	declr(__vbaVarAdd, VarAdd)				\
	declr(__vbaVarAnd, VarAnd)				\
	declr(__vbaVarCat, VarCat)				\
	declr(__vbaVarDiv, VarDiv)				\
	declr(__vbaVarEqv, VarEqv)				\
	declr(__vbaVarIdiv, VarIdiv)			\
	declr(__vbaVarImp, VarImp)				\
	declr(__vbaVarMod, VarMod)				\
	declr(__vbaVarMul, VarMul)				\
	declr(__vbaVarOr, VarOr)				\
	declr(__vbaVarPow, VarPow)				\
	declr(__vbaVarSub, VarSub)				\
	declr(__vbaVarXor, VarXor)				\

#pragma warning (disable : 4477) /* This warning is created by converting types in sprintf */

BUNCH_OF_VAR_MANIPULATIONS(DECLARE_VBA_VARIANT_MANIPULATION_BRIDGE_TO_OLE_MANIPULATION);


/**
 * @brief			Gets the value of a VT_BYREF Variant to another Variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest always.
 */
VARIANTARG * __stdcall VarDerefByref(
	VARIANTARG		*pvargDest,
	VARIANTARG		*pvargSrc
)
{
	// Remove (mask) VT_BYREF
	pvargDest->vt = pvargSrc->vt & ~VT_BYREF;

	switch (pvargDest->vt)
	{
		case VT_I2:
		case VT_BOOL:
		{
			pvargDest->iVal = *pvargSrc->piVal;
			break;
		} /* VT_I2, VT_BOOL */

		case VT_I4:
		case VT_R4:
		case VT_BSTR:
		case VT_ERROR:
		{
			pvargDest->lVal = *pvargSrc->plVal;
			break;
		} /* VT_I4, VT_R4, VT_BSTR, VT_ERROR */

		case VT_R8:
		case VT_CY:
		case VT_DATE:
		{
			pvargDest->cyVal = *pvargSrc->pcyVal;
			break;
		} /* VT_R8, VT_CY, VT_DATE */

		case VT_VARIANT:
		case VT_DECIMAL:
		{
			pvargDest->decVal = *pvargSrc->pdecVal;
			break;
		} /* VT_VARIANT, VT_DECIMAL */

		case VT_UI1:
		{
			pvargDest->bVal = *pvargSrc->pbVal;
			break;
		} /* VT_UI1 */

		case VT_RECORD:
		{
			pvargDest->pvRecord = pvargSrc->pvRecord;
			pvargDest->pRecInfo = pvargSrc->pRecInfo;
			break;
		} /* VT_RECORD */

		case VT_DISPATCH:
		case VT_UNKNOWN:
		{
			pvargDest->plVal = (LONG*)*pvargSrc->plVal;
			break;
		} /* VT_DISPATCH, VT_UNKNOWN */

		default:
		{
			if (!(pvargSrc->vt & VT_ARRAY))
			{
				vbaRaiseException(VBA_EXCEPTION_VARIABLE_USES_A_TYPE_NOT_SUPPORTED_IN_VISUAL_BASIC);
			}

			pvargDest->lVal = *pvargSrc->plVal;
		} /* default */
	} /* switch (pvargSrc->vt) */

	return pvargDest;
} /* VarDerefByref */

/**
 * @brief			Returns a BSTR in a Variant with the hexadecimal representation of a Variant argument.
 * @param			pvargOut		Output variant variable (will be set to VT_BSTR).
 * @param			pvargIn			Input variant to convert.
 */
EXPORT void __stdcall rtcHexVarFromVar(
	VARIANTARG		*pvargOut,
	VARIANTARG		*pvargIn
)
{
	VARIANTARG		vargLocalDeRef;

	/*
		TODO: List of types:
        VT_CY	= 6,
        VT_DATE	= 7,
        VT_BSTR	= 8,
        VT_DISPATCH	= 9,
        VT_ERROR	= 10,
        VT_BOOL	= 11,
        VT_VARIANT	= 12,
        VT_UNKNOWN	= 13,
        VT_DECIMAL	= 14,
        VT_INT	= 22,
        VT_UINT	= 23,
	*/

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	wchar_t lpszwBuffer[20];

	// De-Ref if necessary
	if (pvargIn->vt & VT_BYREF)
	{
		VarDerefByref(&vargLocalDeRef, pvargIn);
		pvargIn = &vargLocalDeRef;
		DEBUG_WIDE(
			"pvargIn->vt & VT_BYREF => deref"
		);
	}

	// For each size of variant, format a different string
	switch (pvargIn->vt)
	{
		case VT_I1:
		case VT_UI1:
		{
			swprintf(lpszwBuffer, 19, L"%.2X", pvargIn->iVal);

			DEBUG_WIDE(
				" 1 byte input '%ls'",
				lpszwBuffer
			);

			break;
		} /* VT_I1, VT_UI1 */

		case VT_I2:
		case VT_UI2:
		{
			swprintf(lpszwBuffer, 19, L"%.4X", pvargIn->iVal);

			DEBUG_WIDE(
				" 2 bytes input '%ls'",
				lpszwBuffer
			);

			break;
		} /* VT_I2, VT_UI2 */

		case VT_I4:
		case VT_R4:
		case VT_UI4:
		{
			swprintf(lpszwBuffer, 19, L"%.8X", pvargIn->lVal);

			DEBUG_WIDE(
				" 4 bytes input '%ls'",
				lpszwBuffer
			);

			break;
		} /* VT_I4, VT_R4, VT_UI4 */

		case VT_R8:
		case VT_I8:
		case VT_UI8:
		{
			swprintf(lpszwBuffer, 19, L"%.16X", pvargIn->ulVal);

			DEBUG_WIDE(
				" 8 bytes input '%ls'",
				lpszwBuffer
			);

			break;
		} /* VT_I8, VT_R8, VT_UI8 */

		default:
		{
			DEBUG_WIDE(
				"vt type not handled %.8x",
				(unsigned int)pvargIn->vt
			);

			swprintf(lpszwBuffer, 19, L"00??");
		} /* default */
	} /* switch (pvargIn->vt) */

	pvargOut->vt = VT_BSTR;
	pvargOut->bstrVal = SysAllocString(lpszwBuffer);

	if (!pvargOut->bstrVal)
	{
		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
	}
} /* rtcHexVarFromVar */

/**
 * @brief			Returns a BSTR in a Variant with the hexadecimal representation of a Variant argument.
 * @param			pvargOut		Output variant variable (will be set to VT_BSTR).
 * @param			pvargIn			Input variant to convert.
 * @returns			pvargOut always.
 */
EXPORT BSTR __stdcall rtcHexBstrFromVar(
	VARIANTARG		*pvargIn
)
{
	VARIANTARG v;
	rtcHexVarFromVar(&v, pvargIn);
	return v.bstrVal;
} /* rtcHexBstrFromVar */

/**
 * @brief			Frees the destination variant and makes a copy of the source variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest on success, 0 on error.
 */
EXPORT VARIANTARG * __fastcall __vbaVarCopy(
	VARIANTARG		*pvargDest,
	VARIANTARG		*pvargSource
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	HRESULT			result = VariantCopy(pvargDest, pvargSource);

	DEBUG_WIDE(
		"pvargDest %.8x, pvargSource %.8x, VariantCopy = %.8x",
		(unsigned int)pvargDest,
		(unsigned int)pvargSource,
		(unsigned int)result
	);

	if (result != S_OK)
	{
		vbaRaiseException(vbaErrorFromHRESULT(result));
		return nullptr;
	}

	return pvargDest;
} /* __vbaVarCopy */

/**
 * @brief			Duplicates the contents of a Variant variable to another Variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest always.
 * @remark			It will try to de-ref the value of VT_BYREF variants. It will call AddRef()
 *					if it's a VT_DISPATCH or VT_UNKNOWN.
 */
EXPORT VARIANTARG * __fastcall __vbaVarDup(
	VARIANTARG		*pvargDest,
	VARIANTARG		*pvargSrc
)
{
	VARIANTARG		vargLocalDeRef;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (pvargSrc->vt & VT_BYREF)
	{
		DEBUG_WIDE(
			"pvargDest %.8x, pvargSrc %.8x, VT_BYREF set, so pvargSrc is now %.8x, pvargSrc->vt %.8x",
			(unsigned int)pvargDest,
			(unsigned int)pvargSrc,
			(unsigned int)&vargLocalDeRef,
			(unsigned int)pvargSrc->vt
		);

		VarDerefByref(&vargLocalDeRef, pvargSrc);
		pvargSrc = &vargLocalDeRef;
	}
	else
	{
		DEBUG_WIDE(
			"pvargDest %.8x, pvargSrc %.8x, pvargSrc->vt %.8x",
			(unsigned int)pvargDest,
			(unsigned int)pvargSrc,
			(unsigned int)pvargSrc->vt
		);
	}

	switch (pvargSrc->vt)
	{
		case VT_DISPATCH:
		case VT_UNKNOWN:
		{
			if (pvargSrc->pdispVal)
			{
				pvargSrc->pdispVal->AddRef();
			}
			// Fall-thru
		} /* VT_DISPATCH, VT_UNKNOWN */

		case VT_EMPTY:
		case VT_NULL:
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DATE:
		case VT_ERROR:
		case VT_BOOL:
		case VT_DECIMAL:
		case VT_UI1:
		case VT_I1:
		{
			memcpy(
				pvargDest,
				pvargSrc,
				sizeof(VARIANTARG)
			);
			break;
		}

		default:
		{
			__vbaVarCopy(pvargDest, pvargSrc);
			break;
		} /* default */
	} /* switch (pvargSrc->vt) */
	
	return pvargDest;
} /* __vbaVarDup */


EXPORT VARIANTARG * __fastcall __vbaVarMove (
	VARIANTARG		*pvargDest,
	VARIANTARG		*pvargSrc
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	/* Free the destination VARIANTARG */
	__vbaFreeVar(pvargDest);

	if (pvargSrc->vt == VT_DISPATCH)
	{
		HRESULT hr = objIDispatchGetDefaultValue(pvargSrc->pdispVal, pvargDest);
		if (hr == S_OK)
		{
			__vbaFreeVar(pvargSrc);
		}
		else
		{
			vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
		}
	}
	else
	{
		__vbaVarDup(pvargDest, pvargSrc);
	}

	return pvargDest;
} /* __vbaVarMove */

/**
 * @brief			Frees a variant variable (including an array)
 * @param			pvargVariant	Pointer to a VARIANTARG that will be freed.
 * @returns			none.
 */
EXPORT void __fastcall __vbaFreeVar(
	VARIANTARG		*pvargVariant
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	if (!pvargVariant)
	{
		DEBUG_ASCII(
			"pvargVariant %.8x -- NULL",
			pvargVariant
		);
		return;
	}

	DEBUG_ASCII(
		"pvargVariant %.8x, pvargVariant->vt %.8x",
		pvargVariant,
		pvargVariant->vt
	);

	if (!(pvargVariant->vt & VT_BYREF))
	{
		switch (pvargVariant->vt & VT_TYPEMASK)
		{
			case VT_BSTR:
			{
				if (pvargVariant->lVal)
				{
					SysFreeString(pvargVariant->bstrVal);
				}
				break;
			} /* VT_BSTR */

			case VT_DISPATCH:
			case VT_UNKNOWN:
			{
				if (pvargVariant->pdispVal)
				{
					pvargVariant->pdispVal->Release();
				}
				break;
			} /* VT_DISPATCH, VT_UNKNOWN */

			default:
			{
				/* Do we have an array? */
				if (!(pvargVariant->vt & VT_ARRAY))
				{
					/** 
					 * Nope, so fall back to the VariantClear
					 * (this is stated in the MSDN if we don't know what to do with it)
					 */
					VariantClear(pvargVariant);
				}
				else
				{
					/* Handle the case for an array */
					if (pvargVariant->pvarVal)
					{
						if (pvargVariant->pvarVal->lVal)
						{
							// TODO: Check this condition
							DEBUG_ASCII(
								"pvargVariant->vt VT ???,  but pvargVariant->pvarVal (%.8x) is not null, and pvarVal->lVal (%.8x) isn't either",
								pvargVariant->pvarVal,
								pvargVariant->pvarVal->lVal
							);

							vbaRaiseException(VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED);
						}
					}

					if (pvargVariant->parray)
					{
						pvargVariant->parray->fFeatures &= 0xFFEFu;
						SafeArrayDestroy(pvargVariant->parray);
					}
				}
			} /* default */
		} /* switch (pvargVariant->vt & 0x7F) */
	}
	else
	{
		DEBUG_ASCII("pvargVariant->vt & VT_BYREF => calling VariantClear");
		VariantClear(pvargVariant);
	}

	pvargVariant->vt = VT_EMPTY;
} /* __vbaFreeVar */

/**
 * @brief			Frees an array of variant variables (including a arrays)
 * @param			dwCount			Count of elements.
 * @param			pvargVariant	Pointer to the first VARIANTARG element that will be freed.
 * @returns			none.
 */
EXPORT void __cdecl __vbaFreeVarList(
	unsigned int	argCount,
	...
)
{
	VARIANTARG	*pvargVariant;

	va_list		args;
	va_start(args, argCount);

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	while (argCount--)
	{
		pvargVariant = va_arg(args, VARIANTARG*);
		DEBUG_WIDE(
			"arg() %.8x, argsRemaining %.8x",
			(unsigned int)pvargVariant,
			argCount
		);

		__vbaFreeVar(pvargVariant);
	}

	va_end(args);
} /* __vbaFreeVarList */