#include "Logging.hpp"
#include "vba_Locale.h"

LCID getUserLocale()
{
    LCID result;

    result = GetUserDefaultLCID();

    if (!result)
    {
        result = 0x0409; /* Default: United States */
    }

    LOG(LOG_DEBUG) << L"locale = " << vbl::Hex((unsigned long)result);

    return result;
}
