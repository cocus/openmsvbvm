#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include "StringManipulation.hpp"

// TODO: Make proper declaration of this
extern void GetVBProjectTitle(BSTR* rhs);

/**
 * @brief			VB's message box implementation.
 * @param			pvargMessage		Message of the dialog box. Can be omitted.
 * @param			uType				Type of the message box. Same as MessageBoxW.
 * @param			pvargTitle			Title to be used in the dialog box. Can be omitted.
 * @param			pvarHelpFile		Help file path. Can be omitted.
 * @param			pvarHelpContext		Help context ID (only if pvarHelpFile is specified). Can be omitted.
 * @returns			An int with the result of the message box (which button was selected).
 */
EXPORT int __stdcall rtcMsgBox(VARIANTARG* pvargMessage, UINT uType, VARIANTARG* pvargTitle, VARIANTARG* pvarHelpFile, VARIANTARG* pvarHelpContext)
{

    LOG(LOG_DEBUG) << L"msg " << vbl::Hex((unsigned long)pvargMessage) << L", uType " << vbl::Hex((unsigned long)uType)
                   << L", title " << vbl::Hex((unsigned long)pvargTitle) << L", a4 " << vbl::Hex((unsigned long)pvarHelpFile)
                   << L", a5 " << vbl::Hex((unsigned long)pvarHelpContext);

    /* Confirmed via real msvbvm60.dll disassembly (ordinal 595): each of the three
       button/icon/default-button sub-fields of uType is range-checked, and if ANY of
       them is out of range, the WHOLE value is silently reset to 0 (vbOKOnly, no icon,
       default button 1) rather than just clamping the offending sub-field. */
    if ((uType & 0xF) > 5 || (uType & 0xF0) > 0x40 || (uType & 0xF00) > 0x300)
    {
        uType = 0;
    }

    BSTR message = __vbaStrErrVarCopy(pvargMessage);
    BSTR title = __vbaStrErrVarCopy(pvargTitle);

    if (!title)
    {
        GetVBProjectTitle(&title);
    }
    if (!message)
    {
        message = SysAllocString(L"");
    }

    // TODO: Whenever forms become available, use the topmost form's handle for the hWnd argument
    int ret = MessageBoxW(0, message, title, uType);

    __vbaFreeStr(&message);
    __vbaFreeStr(&title);

    return ret;
} /* rtcMsgBox */
