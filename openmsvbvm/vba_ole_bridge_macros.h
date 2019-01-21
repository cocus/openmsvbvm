#pragma once



#define DECLARE_VBA_CONVERSION_BRIDGE_TO_OLE_CONVERSION(outType, inType, exportedName, oleAPI, flg, debugInParams)	\
EXPORT outType __stdcall exportedName(																				\
	inType			unkVal																							\
)																													\
{																													\
	HRESULT		result;																								\
	outType		unkLocalCopy;																						\
																													\
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();																			\
																													\
	DEBUG_WIDE(debugInParams, unkVal);																				\
																													\
	result = oleAPI(																								\
		unkVal,																										\
		getUserLocale(),																							\
		flg,																										\
		&unkLocalCopy																								\
	);																												\
																													\
	if (result < 0)																									\
	{																												\
		DEBUG_WIDE("result = %.8x", result);																		\
		vbaRaiseException(vbaErrorFromHRESULT(result));																\
		return NULL;																								\
	}																												\
																													\
	return unkLocalCopy;																							\
}

#define DECLARE_VBA_VARIANT_MANIPULATION_BRIDGE_TO_OLE_MANIPULATION(exportedName, oleAPI)	\
EXPORT LPVARIANT __stdcall exportedName(													\
	LPVARIANT	pvarResult,																	\
	LPVARIANT	pvarRight,																	\
	LPVARIANT	pvarLeft																	\
)																							\
{																							\
	HRESULT		result;																		\
																							\
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();													\
																							\
	DEBUG_WIDE(																				\
		"pvarResult %.8x, pvarRight %.8x, pvarLeft %.8x",									\
		(unsigned int)pvarResult,															\
		(unsigned int)pvarRight,															\
		(unsigned int)pvarLeft																\
	);																						\
																							\
	result = oleAPI(																		\
		pvarLeft,																			\
		pvarRight,																			\
		pvarResult																			\
	);																						\
																							\
	if (result < 0)																			\
	{																						\
		DEBUG_WIDE("result = %.8x", result);												\
		vbaRaiseException(vbaErrorFromHRESULT(result));										\
		return NULL;																		\
	}																						\
																							\
	return pvarResult;																		\
}