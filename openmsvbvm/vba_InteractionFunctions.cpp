#include "vba_internal.h"

#include "vba_exception.h"

#include "vba_strManipulation.h"

#include <strsafe.h>

/**
 * @brief			Sets the focus to the specified window with the corresponding title.
 * @param			pvargTitle		Title of the window.
 * @param			pvargWait		Bool that specifies to wait for user input or not.
 */
EXPORT void __stdcall rtcAppActivate(
	VARIANTARG		* pvargTitle,
	VARIANTARG		* pvargWait
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvargTitle %.8x, pvargWait %.8x",
		(unsigned int)pvargTitle,
		(unsigned int)pvargWait
	);

	UINT			uiWaitTime = 0;
	BSTR			strTitle;

	/* Create a copy of pvargTitle and convert it to BSTR */
	strTitle = __vbaStrErrVarCopy(pvargTitle);

	DEBUG_WIDE(
		"Window title: '%ls', wait time %.8x",
		strTitle,
		uiWaitTime
	);

	HWND hWnd = FindWindowW(nullptr, strTitle);

	SysFreeString(strTitle);

	if (!hWnd)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	DEBUG_WIDE(
		"Window hWnd %.8x",
		(unsigned int)hWnd
	);

	DWORD			dwProcessId;
	DWORD			dwThreadProcessId;

	dwThreadProcessId = GetWindowThreadProcessId(hWnd, &dwProcessId);
	AttachThreadInput(0, dwThreadProcessId, true);

	SetForegroundWindow(hWnd);
	SetFocus(hWnd);

	AttachThreadInput(0, dwThreadProcessId, 0);
} /* rtcAppActivate */

/**
 * @brief			Plays a system beep.
 */
EXPORT void __stdcall rtcBeep()
{
	MessageBeep(MB_OK);
} /* rtcBeep */

/**
 * @brief			Chooses gets an item from a SAFEARRAY using its index.
 * @param			pvargDest		Pointer to a VARIANTARG where the selected element will be copied to.
 * @param			fIndex			Index of selected element within the array. Float because VB is weird.
 * @param			ppsaArray		Pointer to a pointer of a SAFEARRAY that holds the elements to choose from.
 */
EXPORT void __stdcall rtcChoose(
	VARIANTARG			* pvargDest,
	FLOAT				fIndex,
	SAFEARRAY			** ppsaArray)
{
	VARIANTARG			vargLocalCopy;
	VARIANTARG			pv;
	LONG				lIndices = 0;
	HRESULT				hr;

	/* Zero out the destination variant variable */
	memset(
		pvargDest,
		0,
		sizeof(VARIANTARG)
	);

	pvargDest->vt = VT_NULL;

	/* Check for index bounds */
	if (!(fIndex >= 0.0 && fIndex < 2147483600.0))
	{
		// TODO: Raise an exception?
		return;
	}

	/* If the array is more than one dimension, raise an exception */
	if (SafeArrayGetDim(*ppsaArray) != 1)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

#pragma warning (disable : 244) /* We know that we'll loose some data by converting, don't yell at me because of that */
	/* For some weird reason VB passes a float instead of a LONG, so convert that to LONG */
	lIndices = fIndex - 1;
#pragma warning (default : 244) /* Restore previous warning */

	/* Get the element from the array */
	hr = SafeArrayGetElement(
		*ppsaArray,
		&lIndices,
		(void *)&pv
	);

	if (hr < 0)
	{
		vbaRaiseException(vbaErrorFromHRESULT(hr));
		return;
	}

	if (!(pv.vt & VT_BYREF))
	{
		/* VT_BYREF not set */
		memcpy(
			pvargDest,
			&pv,
			sizeof(VARIANTARG)
		);
	}
	else
	{
		/* VT_BYREF set */
		hr = VariantCopyInd(
			&vargLocalCopy,
			&pv
		);

		if (hr < 0)
		{
			vbaRaiseException(vbaErrorFromHRESULT(hr));
			return;
		}
		else
		{
			memcpy(
				pvargDest,
				&vargLocalCopy,
				sizeof(VARIANTARG)
			);
		}
	}

	return;
} /* rtcChoose */

/**
 * @brief			Gets a named environment variable.
 * @param			pvargIn			Environment variable name.
 * @return			BSTR with the choosen environment variable, or ""
 *					if the variable doesn't exists.
 */
EXPORT BSTR __stdcall rtcEnvironBstr(
	VARIANTARG			* pvargIn
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	BSTR		bstrRet;
	BSTR		bstrEnvName;

	bstrEnvName = __vbaStrVarCopy(pvargIn);

	if (!bstrEnvName)
	{
		bstrRet = SysAllocString(L"");
	}
	else
	{
		DEBUG_WIDE(
			"EnvironmentVariable: '%ls'",
			bstrEnvName
		);

		DWORD dwSize;
		dwSize = GetEnvironmentVariableW(
			bstrEnvName,
			NULL,
			0
		);

		if (dwSize == 0)
		{
			/* Specified environment string not found */
			bstrRet = SysAllocString(L"");
		}
		else
		{
			bstrRet = SysAllocStringLen(
				NULL,
				dwSize
			);

			if (!bstrRet)
			{
				vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
			}
			else
			{
				dwSize = GetEnvironmentVariableW(
					bstrEnvName,
					bstrRet,
					dwSize
				);
			}
		}
	}

	return bstrRet;
} /* rtcEnvironBstr */

/**
 * @brief			Gets a named environment variable.
 * @param			pvargOut		Pointer to the destination VARIANTARG.
 * @param			pvargIn			Environment variable name.
 * @return			pvargOut always, with the choosen environment variable, or ""
 *					if the variable doesn't exists.
 */
EXPORT VARIANTARG * __stdcall rtcEnvironVar(
	VARIANTARG			*pvargOut,
	VARIANTARG			*pvargIn
)
{
	pvargOut->vt = VT_BSTR;
	pvargOut->bstrVal = rtcEnvironBstr(pvargIn);

	return pvargOut;
} /* rtcEnvironVar */

/**
 * @brief			Returns a variant between two variant variables based on the
 *					boolean value of a third variable.
 * @param			pvargDest			Pointer to the destination VARIANTARG.
 * @param			pvargBooleanValue	VARIANTARG that selects which value to use.
 * @param			pvargTrueValue		VARIANTARG to return if the boolean value is true.
 * @param			pvargFalseValue		VARIANTARG to return if the boolean value is false.
 * @return			pvargDest always, with a copy of the value of pvargTrueValue
 *					or pvargFalseValue.
 */
EXPORT VARIANTARG * __stdcall rtcImmediateIf(
	VARIANTARG			*pvargDest,
	VARIANTARG			*pvargBooleanValue,
	VARIANTARG			*pvargTrueValue,
	VARIANTARG			*pvargFalseValue
)
{
	HRESULT			hr;
	VARIANTARG		vargB;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	/* Convert pvargBooleanValue to a Boolean variant (deref if necessary) */
	hr = VariantChangeType(
		&vargB,
		pvargBooleanValue,
		0,
		VT_BOOL
	);

	if (hr != S_OK)
	{
		vbaRaiseException(vbaErrorFromHRESULT(hr));
		return pvargDest;
	}

	DEBUG_WIDE(
		"booleanValue %.1d, trueValueVT %.8x, falseValueVT %.8x",
		vargB.boolVal,
		pvargTrueValue->vt,
		pvargFalseValue->vt
	)

		/* Create a copy of pvargTrueValue or pvargFalseValue as needed */
		hr = VariantCopyInd(
			pvargDest,
			vargB.boolVal ? pvargTrueValue : pvargFalseValue
		);

	if (hr != S_OK)
	{
		vbaRaiseException(vbaErrorFromHRESULT(hr));
	}

	return pvargDest;
} /* rtcImmediateIf */

/**
 * @brief			Returns a Variant (String) indicating where a number
 *					occurs within a calculated series of ranges.
 * @param			pvargDest		Pointer to the destination VARIANTARG.
 * @param			pvargNumber		Whole number that you want to evaluate against the ranges.
 * @param			pvargStart		Whole number that is the start of the overall range of numbers. The number can't be less than 0.
 * @param			pvargEnd		Whole number that is the end of the overall range of numbers. The number can't be equal to or less than start.
 * @param			pvargInterval	Whole number that specifies the interval.
 */
EXPORT void __stdcall rtcPartition(
	VARIANTARG			*pvargDest,
	VARIANTARG			*pvargNumber,
	VARIANTARG			*pvargStart,
	VARIANTARG			*pvargEnd,
	VARIANTARG			*pvargInterval
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	VARIANTARG			vargNumber;
	VARIANTARG			vargStart;
	VARIANTARG			vargEnd;
	VARIANTARG			vargInterval;

	int					iRange1, iRange2;

	HRESULT				hr;

	VariantInit(&vargNumber);
	VariantInit(&vargStart);
	VariantInit(&vargEnd);
	VariantInit(&vargInterval);

	/* Convert Number to I4 */
	hr = VariantChangeType(
		&vargNumber,
		pvargNumber,
		0,
		VT_I4
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
		return;
	}

	/* Convert Start to I4 */
	hr = VariantChangeType(
		&vargStart,
		pvargStart,
		0,
		VT_I4
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
		return;
	}

	/* Convert End to I4 */
	hr = VariantChangeType(
		&vargEnd,
		pvargEnd,
		0,
		VT_I4
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
		return;
	}

	/* Convert Interval to I4 */
	hr = VariantChangeType(
		&vargInterval,
		pvargInterval,
		0,
		VT_I4
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
		return;
	}

	DEBUG_WIDE(
		"Number %.d, Start %.d, End %.d, Interval %.d",
		vargNumber.intVal,
		vargStart.intVal,
		vargEnd.intVal,
		vargInterval.intVal
	);

	if (vargInterval.intVal == 1)
	{
		/* If interval is 1, the range is number:number, regardless of the start and stop arguments. */
		iRange1 = vargNumber.intVal;
		iRange2 = iRange1;
	}
	else
	{
		// TODO: Come up with an algorithm for this
	}
} /* rtcPartition */

/**
 * @brief			Sends multiple keys or combination of keys to the focused window.
 * @param			bstrKeys		BSTR with the keys and combination of keys.
 * @param			pvargWaitTime	Pointer to a VARIANTARG with the optional value of
 *									the waiting time after sending the keys.
 */
EXPORT void __stdcall rtcSendKeys(
	BSTR			bstrKeys,
	VARIANTARG		*pvargWaitTime
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"inputKeys '%ls', pvargWaitTime %.8d",
		bstrKeys,
		pvargWaitTime->intVal
	);

	// TODO: implement something like Karl E. Peterson's SendInput
	// http://www.classicvb.net/samples/SendInput/
} /* rtcSendKeys */

/**
 * @brief			Starts up a new process.
 * @param			vargPath		Path to the application.
 * @param			uiShowCmd		Show mode for the newly created app.
 * @return			PID of the newly created process on success, -1 otherwise.
 */
EXPORT DWORD __stdcall rtcShell(
	VARIANTARG		*vargPath,
	unsigned int	uiShowCmd
)
{
	BSTR				bstrCmdLine;
	BOOL				processCreated;
	STARTUPINFOW		StartupInfo;
	PROCESS_INFORMATION	ProcessInformation;

	DWORD				lastError;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"vargPath %.8x, uiShowCmd %.8x",
		(unsigned int)vargPath,
		uiShowCmd
	);

	/* Check if the uiShowCmd argument is valid */
	if (uiShowCmd > 11)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return -1;
	}

	bstrCmdLine = __vbaStrVarCopy(vargPath);
	if (!bstrCmdLine)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL); // Or VBA_EXCEPTION_OUT_OF_STRING_SPACE?
		return -1;
	}

	/* Setup the StartupInfoW struct */
	memset(&StartupInfo, 0, sizeof(StartupInfo));
	StartupInfo.cb = sizeof(StartupInfo);
	StartupInfo.dwFlags = STARTF_USESHOWWINDOW;
	StartupInfo.wShowWindow = uiShowCmd;

	/* Start the process */
	processCreated = CreateProcessW(
		0,
		bstrCmdLine,
		0,
		0,
		0,
		0,
		0,
		0,
		&StartupInfo,
		&ProcessInformation
	);

	lastError = GetLastError();

	SysFreeString(bstrCmdLine);

	if (!processCreated)
	{
		switch (lastError)
		{
			case ERROR_FILE_NOT_FOUND:
			{
				vbaRaiseException(VBA_EXCEPTION_FILE_NOT_FOUND);
				break;
			}
			case ERROR_PATH_NOT_FOUND:
			{
				vbaRaiseException(VBA_EXCEPTION_PATH_NOT_FOUND);
				break;
			}
			case ERROR_OUTOFMEMORY:
			case ERROR_NOT_ENOUGH_MEMORY:
			{
				vbaRaiseException(VBA_EXCEPTION_OUT_OF_MEMORY);
				break;
			}
			default:
			{
				vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
			}
		}

		return -1;
	}

	WaitForInputIdle(ProcessInformation.hProcess, 10000);

	CloseHandle(ProcessInformation.hThread);
	CloseHandle(ProcessInformation.hProcess);

	return ProcessInformation.dwProcessId;
} /* rtcShell */