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

/**
 * @brief			Real implementation behind the VB "Load"/"Unload" statements for a
 *					Form-derived object: creates (Load) or destroys (Unload) its real
 *					Win32 window. object must be one of this project's own wrapped
 *					objects (see TryGetWrapperOf in ObjectManipulation.cpp); anything
 *					else returns E_INVALIDARG.
 */
HRESULT VBFormLoad(
	IDispatch		* object
);

HRESULT VBFormUnload(
	IDispatch		* object
);

/**
 * @brief			Attempts to fire Form_QueryUnload on a Form-derived object's
 *					compiled instance (pVBVTableRaw is its vba_VBVTable*, opaque here)
 *					before its window actually closes. *pCancel is set to 0 up front
 *					and left at whatever the handler set it to on return. Returns
 *					false (with *pCancel unchanged at 0) if no handler could safely be
 *					identified/invoked -- see the scoped heuristic and its limitations
 *					documented on vbFormWrapper::TryFireQueryUnload in
 *					ObjectManipulation.cpp.
 */
bool VBFormTryQueryUnload(
	void			* pVBVTableRaw,
	short			* pCancel
);