#include "vba_internal.h"

#include "vba_exception.h"

#include "vba_strManipulation.h"
#include "vba_objManipulation.h"
#include "vba_varManipulation.h"



/**
 * @brief			Raises the Subscript out of range exception.
 */
EXPORT void __stdcall __vbaGenerateBoundsError()
{
	vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
} /* __vbaGenerateBoundsError */

/**
 * @brief			Locks a SafeArray and copies the pointer to another pointer.
 * @param			ppsaDest		Pointer to a SafeArray pointer, where psaSource will be set to.
 * @param			psaSource		Source SafeArray pointer (to be locked).
 */
EXPORT void __stdcall __vbaAryLock(
	SAFEARRAY		** ppsaDest,
	SAFEARRAY		* psaSource
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"Locking array %.8x",
		(unsigned int)psaSource
	);

	if (!psaSource || !ppsaDest)
	{
		vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
		return;
	}

	if (SafeArrayLock(psaSource) != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED);
		return;
	}

	*ppsaDest = psaSource;
} /* __vbaAryLock */

/**
 * @brief			Unlocks and nulls a pointer to a previously locked SafeArray pointer.
 * @param			ppsaSafeArray	Pointer to a SafeArray pointer.
 */
EXPORT void __stdcall __vbaAryUnlock(
	SAFEARRAY		** ppsaSafeArray
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"Unlocking array %.8x",
		(unsigned int)ppsaSafeArray
	);

	if (!*ppsaSafeArray)
	{
		return;
	}

	if (SafeArrayUnlock(*ppsaSafeArray) != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED);
		return;
	}

	*ppsaSafeArray = 0;
} /* __vbaAryUnlock */

/**
 * @brief			Obtains the lower bound of given dimension of a SafeArray.
 * @param			iDimension		Dimension number (base 1).
 * @param			psaSafeArray	Pointer to the SafeArray.
 * @return			Lower bound of the given dimension, or zero (and a exception) on failure.
 */
EXPORT LONG __stdcall __vbaLbound(
	signed int		iDimension,
	SAFEARRAY		* psaSafeArray
)
{
	if (iDimension < 1)
	{
		vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
		return 0;
	}
	if (!psaSafeArray)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return 0;
	}
	if (iDimension > psaSafeArray->cDims)
	{
		vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
		return 0;
	}

	return psaSafeArray->rgsabound[psaSafeArray->cDims - iDimension].lLbound;
} /* __vbaLbound */

/**
 * @brief			Obtains the upper bound of given dimension of a SafeArray.
 * @param			iDimension		Dimension number (base 1).
 * @param			psaSafeArray	Pointer to the SafeArray.
 * @return			Upper bound of the given dimension, or zero (and a exception) on failure.
 */
EXPORT LONG __stdcall __vbaUbound(
	signed int		iDimension,
	SAFEARRAY		* psaSafeArray
)
{
	if (iDimension < 1)
	{
		vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
		return 0;
	}
	if (!psaSafeArray)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return 0;
	}
	if (iDimension > psaSafeArray->cDims)
	{
		vbaRaiseException(VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE);
		return 0;
	}

	SAFEARRAYBOUND * psabDimensionBounds = &psaSafeArray->rgsabound[psaSafeArray->cDims - iDimension];

	return psabDimensionBounds->lLbound
		 + psabDimensionBounds->cElements;
} /* __vbaUbound */

/**
 * @brief			Creates a SafeArray from a statically defined structure.
 * @param			psaDest			Pointer to the destination SafeArray.
 * @param			psaSource		Pointer to the statically defined SafeArray.
 * @param			lpVoidArg		Pointer to a IRecordInfo, GUID or nothing depending on the
 *									features in the SafeArray.
 */
EXPORT void __stdcall __vbaAryConstruct2(
	SAFEARRAY			* psaDest,
	SAFEARRAY			* psaSource,
	void				* lpVoidArg
)
{
	/* Copy the source array psaSource to the destination array psaDest, including the bounds */
	memcpy(
		psaDest,
		psaSource,
		sizeof(SAFEARRAY) + (psaSource->cDims > 0u ? (sizeof(SAFEARRAYBOUND) * (psaSource->cDims - 1)) : 0)
	);

	/* Is an embedded array in a structure? */
	if (!(psaDest->fFeatures & FADF_EMBEDDED))
	{
		if (SafeArrayAllocData(psaDest) < 0 || !psaDest->pvData)
		{
			SafeArrayDestroyDescriptor(psaDest);
			return;
		}
	}

	/* Do we have an additional argument? */
	if (lpVoidArg)
	{
		/* If this array contains records, then the additional argument is a pointer to a RecordInfo struct */
		if (psaDest->fFeatures & FADF_RECORD)
		{
			psaDest->rgsabound[0].lLbound = 0;
			SafeArraySetRecordInfo(
				psaDest,
				reinterpret_cast<IRecordInfo*>(lpVoidArg)
			);
		}
		/* If this array contains the IID, then the additional argument is a pointer to a GUID struct */
		else if (psaDest->fFeatures & FADF_HAVEIID)
		{
			psaDest->cLocks = 0;
			psaDest->pvData = 0;
			psaDest->rgsabound[0].cElements = 0;
			psaDest->rgsabound[0].lLbound = 0;
			SafeArraySetIID(
				psaDest,
				(const GUID&)lpVoidArg
			);
		}
		else if ((psaSource->fFeatures & FADF_HAVEVARTYPE))
		{
			// ?
		}
	}
} /* __vbaAryConstruct2 */

/**
 * @brief			Erases a SafeArray and nulls the pointer to that SafeArray.
 * @param			pData			?
 * @param			ppSafeArray		Pointer to a SafeArray pointer that will be
 *									destroyed.
 */
EXPORT void __stdcall __vbaErase(
	int				* pData,
	SAFEARRAY		** ppSafeArray
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pData %.8x, ppsaSafeArray %.8x",
		(unsigned int)pData,
		(unsigned int)ppSafeArray
	);

	if (!*ppSafeArray)
	{
		return;
	}

	switch (SafeArrayDestroy(*ppSafeArray))
	{
		case DISP_E_ARRAYISLOCKED:
		{
			vbaRaiseException(VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED);
			break;
		}

		case S_OK:
		{
			*ppSafeArray = nullptr;
			break;
		}

		default:
		{
			vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
			break;
		}
	}
} /* __vbaErase */

/**
 * @brief			Destructs a SafeArray and its contents.
 * @param			pData			?
 * @param			ppSafeArray		Pointer to a SafeArray pointer that will be
 *									destroyed and nulled.
 */
EXPORT void __stdcall __vbaAryDestruct(
	int				* pData,
	SAFEARRAY		** ppSafeArray
)
{
	if (!*ppSafeArray)
	{
		return;
	}

	if (((*ppSafeArray)->fFeatures & FADF_EMBEDDED) ||
		((*ppSafeArray)->fFeatures & FADF_AUTO))
	{
		if ((*ppSafeArray)->fFeatures & FADF_RECORD)
		{
			SafeArraySetRecordInfo(*ppSafeArray, 0);
		}
	}
	else
	{
		if ((*ppSafeArray)->pvData)
		{
			(*ppSafeArray)->fFeatures &= 0xFFEDu;
			SafeArrayDestroyData(*ppSafeArray);

			if ((*ppSafeArray)->fFeatures & FADF_RECORD)
			{
				SafeArraySetRecordInfo(*ppSafeArray, 0);
			}
		}

		// This might be sufficient, since SafeArrayDestroy also
		// does the same as SafeArrayDestroyData.
		__vbaErase(pData, ppSafeArray);
	}
} /* __vbaAryDestruct */

/**
 * @brief			Redims a given dimension of a SafeArray, preserving it's contents.
 * @param			ppsaSafeArray	Pointer to a SafeArray pointer.
 * @param			uiDimensions	Number of dimensions to redim.
 * @param			...				New value for each dimension.
 */
EXPORT void __cdecl __vbaRedimPreserve(
	int				iUnk1,
	int				iUnk2,
	SAFEARRAY		** ppsaSafeArray,
	int				iUnk3,
	unsigned int	uiDimensions,
	...
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ppsaSafeArray %.8x, uiDimensions %.8x",
		(unsigned int)ppsaSafeArray,
		uiDimensions
	);

	va_list		args;
	va_start(args, uiDimensions);

	if (uiDimensions == 1)
	{
		/* Redim a one-dimension array by just calling SafeArrayRedim */
		int iNewElementCount = va_arg(args, int);

		DEBUG_WIDE(
			"iNewElementCount = %d, PreviousElementCount %d",
			iNewElementCount,
			(*ppsaSafeArray)->rgsabound[0].cElements
		);

		SAFEARRAYBOUND sabNewBound;
		memcpy(
			&sabNewBound,
			&(*ppsaSafeArray)->rgsabound[0],
			sizeof(sabNewBound)
		);
		sabNewBound.cElements = iNewElementCount;

		HRESULT hr = SafeArrayRedim(
			*ppsaSafeArray,
			&sabNewBound
		);
	}
	else
	{
		// TODO: Multi dimension array redim!
	}

	va_end(args);
} /* __vbaRedimPreserve */

/**
 * @brief			Redims a given dimension of a SafeArray, nuking it's contents.
 * @param			uFeatures		SafeArray features.
 * @param			ulElementsSize	Size of each element of the array.
 * @param			ppsaOut			Pointer to a SafeArray pointer.
 * @param			uiDimensions	Number of new dimensions to redim.
 * @param			...				Lower and Upper bound for each dimension (two different arguments per dimension).
 */
EXPORT void __cdecl __vbaRedim(
	USHORT			uFeatures,
	ULONG			ulElementsSize,
	SAFEARRAY		** ppsaOut,
	void			* lpArgument,
	USHORT			uiDimensions,
	...
)
{
	if (!ppsaOut)
	{
		// Wtf?
		return;
	}

	if (*ppsaOut)
	{
		// We have a previously defined array, so nuke it first
		__vbaErase((int*)lpArgument, ppsaOut);
	}

	HRESULT			hr;

	if (uFeatures & (FADF_RECORD | FADF_HAVEIID | FADF_HAVEVARTYPE))
	{
		// If the array's feature is RECORD, HAVEIID or HAVEVARTYPE
		// then create the descriptor of the SafeArray using SafeArrayAllocDescriptorEx
		
		VARTYPE			vtArrayVarType;
		
		if (uFeatures & FADF_HAVEIID)
		{
			vtArrayVarType = VT_DISPATCH; // ?
		}
		else if (uFeatures & FADF_RECORD)
		{
			vtArrayVarType = VT_RECORD;
		}
		else
		{
			vtArrayVarType = (VARTYPE)lpArgument; // ?
		}

		// Allocate the descriptor for this SafeArray with given vartype
		hr = SafeArrayAllocDescriptorEx(
			vtArrayVarType,
			uiDimensions,
			ppsaOut
		);

		if (hr != S_OK)
		{
			vbaRaiseException(vbaErrorFromHRESULT(hr));
			return;
		}

		if (uFeatures & FADF_HAVEIID)
		{
			// If we have a IID in the lpArgument, then set it via SafeArraySetIID
			hr = SafeArraySetIID(
				*ppsaOut,
				reinterpret_cast<const GUID&>(lpArgument)
			);

			if (hr != S_OK)
			{
				SafeArrayDestroyDescriptor(*ppsaOut);

				vbaRaiseException(vbaErrorFromHRESULT(hr));
				return;
			}
		}
		else if (uFeatures & FADF_RECORD)
		{
			// If we have a IRecordInfo* in the lpArgument, then set it via SafeArraySetRecordInfo
			hr = SafeArraySetRecordInfo(
				*ppsaOut,
				reinterpret_cast<IRecordInfo*>(lpArgument)
			);

			if (hr != S_OK)
			{
				SafeArrayDestroyDescriptor(*ppsaOut);

				vbaRaiseException(vbaErrorFromHRESULT(hr));
				return;
			}
		}	
	}
	else
	{
		// Otherwise, create a standard SafeArray
		hr = SafeArrayAllocDescriptor(
			uiDimensions,
			ppsaOut
		);

		if (hr != S_OK)
		{
			vbaRaiseException(vbaErrorFromHRESULT(hr));
			return;
		}
	}

	// Set the members of the SafeArray with given arguments
	(*ppsaOut)->fFeatures = uFeatures;
	(*ppsaOut)->cbElements = ulElementsSize;

	va_list		args;
	va_start(args, uiDimensions);

	// Set the bounds of the dimensions
	int iCurrDim = 0;
	while (uiDimensions--)
	{
		// We have two elements in the stack per each dimension:
		// First one is the Lower-Bound
		// Second one is the Element count (i.e. Upper-Bound)
		(*ppsaOut)->rgsabound[iCurrDim].cElements = va_arg(args, int) + 1;
		(*ppsaOut)->rgsabound[iCurrDim].lLbound = va_arg(args, int);

		// Fix the element count by subtracting the lower bound
		(*ppsaOut)->rgsabound[iCurrDim].cElements -= (*ppsaOut)->rgsabound[iCurrDim].lLbound;

		iCurrDim++;
	}

	va_end(args);

	// Now alloc the data for all the dimensions and elements
	hr = SafeArrayAllocData(*ppsaOut);

	if (hr != S_OK)
	{
		SafeArrayDestroyDescriptor(*ppsaOut);
		*ppsaOut = 0;

		vbaRaiseException(vbaErrorFromHRESULT(hr));
	}
} /* __vbaRedim */