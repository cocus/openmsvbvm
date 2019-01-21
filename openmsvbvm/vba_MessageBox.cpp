#include "vba_internal.h"
#include "vba_exception.h"

#include "vba_strManipulation.h"

/**
 * @brief			VB's message box implementation.
 * @param			pvargMessage		Message of the dialog box. Can be omitted.
 * @param			uType				Type of the message box. Same as MessageBoxW.
 * @param			pvargTitle			Title to be used in the dialog box. Can be omitted.
 * @param			pvarHelpFile		Help file path. Can be omitted.
 * @param			pvarHelpContext		Help context ID (only if pvarHelpFile is specified). Can be omitted.
 * @returns			An int with the result of the message box (which button was selected).
 */
EXPORT int __stdcall rtcMsgBox(
	VARIANTARG		*pvargMessage,
	UINT			uType,
	VARIANTARG		*pvargTitle,
	VARIANTARG		*pvarHelpFile,
	VARIANTARG		*pvarHelpContext
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"msg %.8x, uType %.8x, title %.8x, a4 %.8x, a5 %.8x",
		(unsigned int)pvargMessage,
		uType,
		(unsigned int)pvargTitle,
		(unsigned int)pvarHelpFile,
		(unsigned int)pvarHelpContext
	);

	BSTR message = __vbaStrErrVarCopy(pvargMessage);
	BSTR title = __vbaStrErrVarCopy(pvargTitle);

	if (!title)
	{
		title = SysAllocString(L"<openmsvbvm.dll>");
	}
	if (!message)
	{
		message = SysAllocString(L"");
	}

	// TODO: Whenever forms become available, use the topmost form's handle for the hWnd argument
	int ret = MessageBoxW(
		0,
		message,
		title,
		uType
	);

	__vbaFreeStr(&message);
	__vbaFreeStr(&title);

	return ret;
} /* rtcMsgBox */