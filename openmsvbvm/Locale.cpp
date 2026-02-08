#include "vba_Locale.h"

LCID getUserLocale()
{
	LCID result;

	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	result = GetUserDefaultLCID();

	if (!result)
	{
		result = 0x0409;	/* Default: United States */
	}

	DEBUG_ASCII(
		"locale = %.8x",
		(unsigned int)result
	);

	return result;
}