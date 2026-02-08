#pragma once
#include "vba_internal.h"

/**
* @brief			Frees a list of COM Objects (IUnknowns) via their pointers, and nulls them.
* @param			argCount		Count of elements.
* @param			...				Pointers to objects to free and null them.
*/
EXPORT void __cdecl __vbaFreeObjList(
	unsigned int argCount,
	...
);

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT void __fastcall __vbaFreeObj(
	IUnknown	** punkObj
);

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
EXPORT VARIANTARG * __cdecl __vbaVarLateMemCallLd(
	VARIANTARG		* pvargRet,
	VARIANTARG		* pvarObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
);

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
EXPORT VARIANTARG * __cdecl __vbaVarLateMemCallLdRf(
	VARIANTARG		* pvargRet,
	VARIANTARG		* pvarObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
);
/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
EXPORT void __cdecl __vbaLateMemCall(
	IDispatch		* pidObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
);
/**
 * @brief			Returns a pointer to an IDispatch from a VARIANTARG.
 * @param			pvargIn			Pointer to the VARIANTARG where the IDispatch will be extracted from.
 * @returns			plVal from the VARIANTARG argument, only if the object type is VT_DISPATCH.
 */
EXPORT IDispatch * __stdcall __vbaObjVar(
	VARIANTARG * pvargIn
);
/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall rtcCreateObject2(
	VARIANTARG *pvargObject,
	BSTR bstrClassName,
	BSTR bstrServerName
);

HRESULT objIDispatchGetDefaultValue(
	IDispatch		* pidObject,
	VARIANTARG		* pvargValueOut
);