#include "vba_internal.h"
#include "vba_exception.h"

/**
 * @brief			Compares two BSTRs using OLE's VarBstrCmp .
 * @param			compare_method	Compare method (vbBinaryCompare, vbTextCompare, etc)
 * @param			bstrRight		First string.
 * @param			bstrLeft		Second string.
 * @returns			Returns -1 if bstrLeft is less than bstrRight, 0 if they're equal, and 1 if bstrLeft is greater than bstrRight.
 */
EXPORT int __stdcall __vbaStrComp(
	int			compare_method,
	BSTR		bstrRight,
	BSTR		bstrLeft
)
{
	HRESULT		result;
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (bstrRight && bstrLeft)
	{
		DEBUG_WIDE(
			"compare_method, %.8x, bstrRight %.8x '%ls', bstrLeft %.8x '%ls'",
			(unsigned int)compare_method,
			(unsigned int)bstrRight,
			bstrRight,
			(unsigned int)bstrLeft,
			bstrLeft
		);
	}
	else
	{
		DEBUG_WIDE(
			"compare_method, %.8x, bstrRight %.8x, bstrLeft %.8x",
			(unsigned int)compare_method,
			(unsigned int)bstrRight,
			(unsigned int)bstrLeft
		);
	}

	if (compare_method == NORM_IGNORENONSPACE)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return 0;
	}

	result = VarBstrCmp(
		bstrLeft,
		bstrRight,
		compare_method,
		0x30001
	);

	if (result >= 0)
	{
		/* Convert the VARCMP_ values to -1, 0 and 1 (as VB expects) */
		return result - 1;
	}

	vbaRaiseException(vbaErrorFromHRESULT(result));

	return 0;
} /* __vbaStrComp */