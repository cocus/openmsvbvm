#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

// TODO: Make this thread-dependant?
static BSTR m_bstrCommandLine;

/**
 * @brief			Returns a BSTR with the app command line.
 * @returns			A BSTR with the full command line.
 */
EXPORT BSTR __stdcall rtcCommandBstr(void)
{
    BSTR result;

    if (m_bstrCommandLine)
    {
        result = SysAllocString(m_bstrCommandLine);
    }
    else
    {
        result = SysAllocString(L"");
    }

    LOG(LOG_DEBUG) << L"psz '" << vbl::Bstr(result) << L"'";

    return result;
}

/**
 * @brief			Returns a Variant with the app command line.
 * @param			pvargDest		Destination Variant variable.
 * @returns			A Variant (VT_BSTR) with the full command line.
 */
EXPORT VARIANTARG* __stdcall rtcCommandVar(VARIANTARG* pvargDest)
{
    if (pvargDest)
    {
        pvargDest->vt = VT_BSTR;
        pvargDest->bstrVal = rtcCommandBstr();
    }

    return pvargDest;
}

OLECHAR* __stdcall rtcSetCommandLine(OLECHAR* psz)
{

    LOG(LOG_DEBUG) << L"psz '" << vbl::Bstr(psz) << L"'";

    if (m_bstrCommandLine)
    {
        SysFreeString(m_bstrCommandLine);
    }

    m_bstrCommandLine = SysAllocString(psz);

    return m_bstrCommandLine;
}
