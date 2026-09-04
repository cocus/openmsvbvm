#pragma once



#define DECLARE_VBA_CONVERSION_BRIDGE_TO_OLE_CONVERSION(outType, inType, exportedName, oleAPI, flg)	\
EXPORT outType __stdcall exportedName(																				\
	inType			unkVal																							\
)																													\
{																													\
	HRESULT		result;																								\
	outType		unkLocalCopy;																						\
																													\
																													\
	LOG(LOG_DEBUG) << L"in '" << unkVal << L"'";																				\
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
		LOG(LOG_DEBUG) << L"result = " << vbl::Hex((unsigned long)result);																		\
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
																							\
	LOG(LOG_DEBUG) << L"pvarResult " << vbl::Hex((unsigned long)pvarResult)					\
		<< L", pvarRight " << vbl::Hex((unsigned long)pvarRight)								\
		<< L", pvarLeft " << vbl::Hex((unsigned long)pvarLeft);								\
																							\
	result = oleAPI(																		\
		pvarLeft,																			\
		pvarRight,																			\
		pvarResult																			\
	);																						\
																							\
	if (result < 0)																			\
	{																						\
		LOG(LOG_DEBUG) << L"result = " << vbl::Hex((unsigned long)result);												\
		vbaRaiseException(vbaErrorFromHRESULT(result));										\
		return NULL;																		\
	}																						\
																							\
	return pvarResult;																		\
}