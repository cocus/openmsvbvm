#pragma once
#include "vba_internal.h"

EXPORT void __cdecl __vbaFreeVarList(
	unsigned int argCount,
	...
);

EXPORT void __fastcall __vbaFreeVar(
	VARIANTARG *pvargVariant
);

EXPORT VARIANTARG * __fastcall __vbaVarDup(
	VARIANTARG *pvarg,
	VARIANTARG *pvargSrc
);

EXPORT VARIANTARG * __fastcall __vbaVarCopy(VARIANTARG *pvarDest, VARIANTARG *pvargSource);

VARIANTARG * __stdcall VarDerefByref(
	VARIANTARG		*pvargDest,
	VARIANTARG		*pvargSrc
);