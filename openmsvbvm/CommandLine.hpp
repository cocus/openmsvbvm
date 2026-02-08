#pragma once

#include "vba_internal.h"

OLECHAR * __stdcall rtcSetCommandLine(
	OLECHAR *psz
);

EXPORT BSTR __stdcall rtcCommandBstr(void);