#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"
#include "vba_Locale.h"

/**
 * @brief			Compares two BSTRs using OLE's VarBstrCmp .
 * @param			compare_method	Compare method (vbBinaryCompare, vbTextCompare, etc)
 * @param			bstrRight		First string.
 * @param			bstrLeft		Second string.
 * @returns			Returns -1 if bstrLeft is less than bstrRight, 0 if they're equal, and 1 if bstrLeft is greater than bstrRight.
 */
EXPORT int __stdcall __vbaStrComp(int compare_method, BSTR bstrRight, BSTR bstrLeft)
{
    HRESULT result;

    if (bstrRight && bstrLeft)
    {
        LOG(LOG_DEBUG) << L"compare_method, " << vbl::Hex((unsigned long)compare_method) << L", bstrRight "
                       << vbl::Hex((unsigned long)bstrRight) << L" '" << vbl::Bstr(bstrRight) << L"', bstrLeft "
                       << vbl::Hex((unsigned long)bstrLeft) << L" '" << vbl::Bstr(bstrLeft) << L"'";
    }
    else
    {
        LOG(LOG_DEBUG) << L"compare_method, " << vbl::Hex((unsigned long)compare_method) << L", bstrRight "
                       << vbl::Hex((unsigned long)bstrRight) << L", bstrLeft " << vbl::Hex((unsigned long)bstrLeft);
    }

    if (compare_method == NORM_IGNORENONSPACE)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return 0;
    }

    result = VarBstrCmp(bstrLeft, bstrRight, getUserLocale(), compare_method);

    if (result >= 0)
    {
        /* Convert the VARCMP_ values to -1, 0 and 1 (as VB expects) */
        return result - 1;
    }

    vbaRaiseException(vbaErrorFromHRESULT(result));

    return 0;
} /* __vbaStrComp */

/**
 * @brief			Compares two BSTRs using binary comparison (vbBinaryCompare).
 *					Confirmed via real msvbvm60.dll disassembly (ordinal 406) to be a
 *					thin forwarder to __vbaStrComp with compare_method hardcoded to 0.
 * @param			bstrRight		First string.
 * @param			bstrLeft		Second string.
 * @returns			-1 if bstrLeft is less than bstrRight, 0 if equal, 1 if greater.
 */
EXPORT int __stdcall __vbaStrCmp(BSTR bstrRight, BSTR bstrLeft)
{
    return __vbaStrComp(0, bstrRight, bstrLeft);
} /* __vbaStrCmp */
