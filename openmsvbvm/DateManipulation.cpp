#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include "vba_Locale.h"

#include <time.h>

/**
 * @brief			Gets the current date time into a VARIANTARG.
 * @param			pvargResult		Pointer to the destination VARIANTARG.
 */
EXPORT void __stdcall rtcGetPresentDate(VARIANTARG* pvargResult)
{
    UDATE udDate;
    HRESULT result;

    GetLocalTime(&udDate.st);
    // TODO: Do this in a nicer way!
    time_t rawtime;
    time(&rawtime);
    struct tm t;
    localtime_s(&t, &rawtime);
    char buff[30];
    strftime(buff, sizeof(buff), "%j", &t);
    udDate.wDayOfYear = atoi(buff);

    pvargResult->vt = VT_DATE;

    result = VarDateFromUdate(&udDate, VAR_VALIDDATE, &pvargResult->date);

    LOG(LOG_DEBUG) << L"pvargResult " << vbl::Hex((unsigned long)pvargResult) << L", pvargResult->vt "
                   << vbl::Hex((unsigned long)pvargResult->vt) << L", pvargResult->date "
                   << vbl::Hex((unsigned long)pvargResult->date, 16) << L", VarDateFromUdate = " << vbl::Hres(result);

    if (result != S_OK)
    {
        vbaRaiseException(vbaErrorFromHRESULT(result));
        return;
    }
} /* rtcGetPresentDate */

/**
 * @brief			Converts a VARIANTARG to DATE.
 * @param			pvargDate		Pointer to the source VARIANTARG.
 * @return			DATE representation of pvargDate, and 0.0 if fails to convert.
 */
EXPORT DATE __stdcall __vbaDateVar(VARIANTARG* pvargDate)
{
    LOG(LOG_DEBUG) << L"pvargDate " << vbl::Hex((unsigned long)pvargDate);

    if (pvargDate->vt == VT_DATE)
    {
        // Don't convert it since it's already a VT_DATE
        return pvargDate->date;
    }
    else
    {
        VARIANTARG vargDateTemp;
        HRESULT hr;

        // Convert to VT_DATE
        hr = VariantChangeType(&vargDateTemp, pvargDate, 0, VT_DATE);

        if (hr != S_OK)
        {
            vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
            return 0.0;
        }

        return vargDateTemp.date;
    }
} /* __vbaDateVar */
