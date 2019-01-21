#include "vba_internal.h"
#include "vba_exception.h"
#include "vba_structures.h"

/**
 * @brief			Loads a DLL and gets the address of a specified procedure.
 * @param			dllTemplate		Pointer to a serDllTemplate structure, with the proper information
 *									about the DLL and procedure to call.
 * @return			FARPROC (func pointer) of the specified procedure on success, 0 otherwise.
 */
EXPORT FARPROC __stdcall DllFunctionCall(
	struct serDllTemplate * dllTemplate
)
{
	FARPROC			hFuncAddr;
	HMODULE			hModule;

	if (!dllTemplate)
	{
		vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
	}

	hModule = LoadLibraryA(
		dllTemplate->lpLibraryNameA
	);
	if (!hModule)
	{
		RaiseExceptionIfLastErrorIsSet();
	}

	hFuncAddr = GetProcAddress(
		hModule,
		dllTemplate->lpProcAddressA
	);
	if (!hFuncAddr)
	{
		RaiseExceptionIfLastErrorIsSet();
	}

	return hFuncAddr;
} /* DllFunctionCall */