#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include <cstring>
#include <string>

#include "ObjectManipulation.hpp"
#include "VariantManipulation.hpp"

/**
 * @brief			Gets the string length of a BSTR.
 * @param			bstrIn				Input string.
 * @returns			Size of the string if the BSTR is valid, 0 otherwise.
 */
UINT __stdcall strSafeGetLength(BSTR bstrIn)
{
    if (bstrIn)
    {
        // The string length is stored in the previous 4 bytes
        // of the starting character of the BSTR. It's weird.
        uint32_t* puintSize = (uint32_t*)bstrIn - 1;

        // Check if we're going to create an exception if we read the size.
        if (IsBadReadPtr(puintSize, sizeof(uint32_t)))
        {
            return 0;
        }

        // Size is in bytes, not in wide-chars, so divide by two.
        return *puintSize / 2;
    }

    return 0;
} /* strSafeGetLength */

/**
 * @brief			Gets the length (in characters) of a BSTR.
 * @param			bstrIn			Input string.
 * @returns			Length in characters, or 0 if bstrIn is null.
 */
EXPORT int __stdcall __vbaLenBstr(BSTR bstrIn)
{
    int iLen = (int)strSafeGetLength(bstrIn);

    LOG(LOG_TRACE) << L"bstrIn " << vbl::Hex((unsigned long)bstrIn) << L", len " << vbl::Hex((unsigned long)iLen);

    return iLen;
} /* __vbaLenBstr */

/**
 * @brief			Implements Asc$().
 * @param			bstrIn			Input string.
 * @returns			ANSI character code of the first character.
 */
EXPORT short __stdcall rtcAnsiValueBstr(BSTR bstrIn)
{

    if (!bstrIn)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return 0;
    }

    UINT uiLen = strSafeGetLength(bstrIn);

    LOG(LOG_TRACE) << L"bstrIn " << vbl::Hex((unsigned long)bstrIn) << L" '" << vbl::Bstr(bstrIn) << L"', len "
                   << vbl::Hex((unsigned long)uiLen);

    if (uiLen == 0)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return 0;
    }

    char szFirstChar[2] = {0};

    /* TODO: The real DLL takes a separate path when this conversion produces 2 bytes
       (a DBCS lead+trail byte pair, e.g. East Asian code pages) that isn't replicated
       here yet -- this always returns just the single ANSI byte. */
    int iChars = WideCharToMultiByte(CP_ACP, 0, (LPCWCH)bstrIn, 1, szFirstChar, sizeof(szFirstChar), 0, 0);

    if (iChars <= 0)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return 0;
    }

    return (unsigned char)szFirstChar[0];
} /* rtcAnsiValueBstr */

/**
 * @brief			Implements Mid$(string, start, [length]) into a Variant result slot.
 * @param			pvargResult		Destination Variant, receives the VT_BSTR result.
 * @param			pvargString		Source string Variant (BSTR or BYREF BSTR).
 * @param			iStart			1-based start position.
 * @param			pvargLength		Length Variant, or a VT_ERROR Variant if Length
 *									was omitted (meaning "to the end of the string").
 * @returns			pvargResult.
 */
EXPORT VARIANTARG* __stdcall rtcMidCharVar(VARIANTARG* pvargResult, VARIANTARG* pvargString, int iStart, VARIANTARG* pvargLength)
{

    LOG(LOG_TRACE) << L"pvargResult " << vbl::Hex((unsigned long)pvargResult) << L", pvargString "
                   << vbl::Hex((unsigned long)pvargString) << L", iStart " << vbl::Hex((unsigned long)iStart) << L", pvargLength "
                   << vbl::Hex((unsigned long)pvargLength);

    if (iStart < 1)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return pvargResult;
    }

    if (!pvargString)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return pvargResult;
    }

    VARIANT vargStr = *pvargString;

    if (pvargString->vt & VT_BYREF)
    {
        VarDerefByref(&vargStr, pvargString);
    }

    if (vargStr.vt != VT_BSTR)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return pvargResult;
    }

    BSTR bstrSrc = vargStr.bstrVal;
    UINT uiLen = strSafeGetLength(bstrSrc);

    /* -1 means "no Length given -- take up to the end of the string" */
    int iLength = -1;

    if (pvargLength && pvargLength->vt != VT_ERROR)
    {
        VARIANT vargLen;
        VariantInit(&vargLen);

        if (FAILED(VariantChangeType(&vargLen, pvargLength, 0, VT_I4)))
        {
            vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
            return pvargResult;
        }

        iLength = vargLen.lVal;

        if (iLength < 0)
        {
            vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
            return pvargResult;
        }
    }

    BSTR bstrResult;

    if ((UINT)(iStart - 1) >= uiLen)
    {
        bstrResult = SysAllocString(L"");
    }
    else
    {
        UINT uiAvailable = uiLen - (iStart - 1);
        UINT uiTake = (iLength < 0 || (UINT)iLength > uiAvailable) ? uiAvailable : (UINT)iLength;

        bstrResult = SysAllocStringLen(bstrSrc + (iStart - 1), uiTake);
    }

    if (!bstrResult)
    {
        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
        return pvargResult;
    }

    if (pvargResult)
    {
        pvargResult->vt = VT_BSTR;
        pvargResult->bstrVal = bstrResult;
    }

    return pvargResult;
} /* rtcMidCharVar */

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcReplace(BSTR bstrExpression, BSTR bstrFind, BSTR bstrReplace, int iStart, int iCount, int iCompareMethod)
{

    LOG(LOG_TRACE) << L"bstrExpression " << vbl::Hex((unsigned long)bstrExpression) << L" '" << vbl::Bstr(bstrExpression)
                   << L"', bstrFind " << vbl::Hex((unsigned long)bstrFind) << L" '" << vbl::Bstr(bstrFind) << L"', bstrReplace "
                   << vbl::Hex((unsigned long)bstrReplace) << L" '" << vbl::Bstr(bstrReplace) << L"', iStart "
                   << vbl::Hex((unsigned long)iStart) << L", iCount " << vbl::Hex((unsigned long)iCount) << L", iCompareMethod "
                   << vbl::Hex((unsigned long)iCompareMethod);
    return bstrExpression;
} /* rtcReplace */

/**
 * @brief			Copies a BSTR and converts their characters to uppercase.
 * @param			strIn			BSTR String to convert.
 * @returns			Converted BSTR, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcUpperCaseBstr(BSTR strIn)
{
    UINT uStrSize = strSafeGetLength(strIn);
    BSTR bstrRet;

    LOG(LOG_TRACE) << L"strIn " << vbl::Hex((unsigned long)strIn) << L", size " << vbl::Hex((unsigned long)uStrSize);

    bstrRet = SysAllocStringLen(strIn, uStrSize);

    if (!bstrRet)
    {
        LOG(LOG_TRACE) << L"SysAllocStringLen failed, err = " << vbl::Hex((unsigned long)GetLastError());

        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
        return nullptr;
    }

    CharUpperBuffW(bstrRet, uStrSize + 1);

    LOG(LOG_TRACE) << L"ret = '" << vbl::Bstr(bstrRet) << L"'";

    return bstrRet;
} /* rtcUpperCaseBstr */

/**
 * @brief			Copies a BSTR and converts their characters to lowercase.
 * @param			strIn			String to convert.
 * @returns			Converted BSTR, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall rtcLowerCaseBstr(BSTR strIn)
{
    UINT uStrSize = strSafeGetLength(strIn);
    BSTR bstrRet;

    LOG(LOG_TRACE) << L"strIn " << vbl::Hex((unsigned long)strIn) << L", size " << vbl::Hex((unsigned long)uStrSize);

    bstrRet = SysAllocStringLen(strIn, uStrSize);

    if (!bstrRet)
    {
        LOG(LOG_TRACE) << L"SysAllocStringLen failed, err = " << vbl::Hex((unsigned long)GetLastError());

        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
        return nullptr;
    }

    CharLowerBuffW(bstrRet, uStrSize + 1);

    LOG(LOG_TRACE) << L"ret = '" << vbl::Bstr(bstrRet) << L"'";

    return bstrRet;
} /* rtcLowerCaseBstr */

/**
 * @brief			Concatenates two BSTRs.
 * @param			bstrRight		Right operand for the concatenation.
 * @param			bstrLeft		Left operand for the concatenation.
 * @returns			Concatenated BSTR on success, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrCat(BSTR bstrLeft, BSTR bstrRight)
{
    if (bstrRight && bstrLeft)
    {
        LOG(LOG_TRACE) << L"bstrRight " << vbl::Hex(bstrRight) << L" '" << vbl::Bstr(bstrRight) << L"', bstrLeft "
                       << vbl::Hex(bstrLeft) << L" '" << vbl::Bstr(bstrLeft) << L"'";
    }
    else
    {
        LOG(LOG_TRACE) << L"bstrRight " << vbl::Hex(bstrRight) << L", bstrLeft " << vbl::Hex(bstrLeft);
    }

    HRESULT result;
    BSTR ret;

    result = VarBstrCat(bstrRight, bstrLeft, &ret);

    if (result < 0)
    {
        LOG(LOG_WARN) << L"result " << vbl::Hres(result);
        vbaRaiseException(vbaErrorFromHRESULT(result));
        return nullptr;
    }

    LOG(LOG_TRACE) << L"result " << vbl::Hres(result) << L", '" << vbl::Bstr(ret) << L"'";

    return ret;
} /* __vbaStrCat */

/**
 * @brief			Moves one BSTR to the pointer of another (previously freed if needed) BSTR.
 * @param			pbstrDest		Where the source BSTR value will be copied to. Freed if not zero.
 * @param			bstrSrc			Source BSTR.
 * @returns			Same value as pbstrSrc.
 */
EXPORT BSTR __fastcall __vbaStrMove(BSTR* pbstrDest, BSTR bstrSrc)
{

    LOG(LOG_TRACE) << L"pbstrDest " << vbl::Hex((unsigned long)pbstrDest) << L", bstrSrc " << vbl::Hex((unsigned long)bstrSrc);

    if (*pbstrDest)
    {
        LOG(LOG_TRACE) << L"pbstrDest contained '" << vbl::Bstr(*pbstrDest) << L"'";

        SysFreeString(*pbstrDest);
    }

    if (bstrSrc)
    {
        LOG(LOG_TRACE) << L"bstrSrc '" << vbl::Bstr(bstrSrc) << L"'";
    }

    *pbstrDest = bstrSrc;
    return bstrSrc;
} /* __vbaStrMove */

/**
 * @brief			Converts a Variant (including VT_ERROR) to a BSTR using OLE's VariantChangeType.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrErrVarCopy(VARIANTARG* pvargIn)
{

    VARIANT vargRet;
    VariantInit(&vargRet);

    if (!pvargIn)
    {
        return nullptr;
    }

    LOG(LOG_TRACE) << L"pvargIn " << vbl::Hex((unsigned long)pvargIn) << L", pvargIn->vt " << vbl::Hex((unsigned long)pvargIn->vt);

    HRESULT hr = VariantChangeType(&vargRet, pvargIn, VARIANT_ALPHABOOL | VARIANT_LOCALBOOL, VT_BSTR);

    if (hr != S_OK)
    {
        return nullptr;
    }

    UINT uStrSize = strSafeGetLength(vargRet.bstrVal);

    LOG(LOG_TRACE) << L"vargRet.bstrVal " << vbl::Hex((unsigned long)vargRet.bstrVal) << L", size " << vbl::Hex((unsigned long)uStrSize);

    BSTR bstrRet = SysAllocStringLen(vargRet.bstrVal, uStrSize);

    __vbaFreeVar(&vargRet);

    if (!bstrRet)
    {
        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
    }

    return bstrRet;
} /* __vbaStrErrVarCopy */

/**
 * @brief			Copies one BSTR to the pointer of another (previously freed if needed) BSTR.
 * @param			strOut			Where the source BSTR value will be copied to. Freed if not zero.
 * @param			strIn			Source BSTR.
 * @returns			Contents of strOut.
 */
EXPORT BSTR __fastcall __vbaStrCopy(BSTR* strOut, BSTR strIn)
{
    BSTR bstrRet = nullptr;
    UINT uStrSize = 0;

    if (strIn)
    {
        uStrSize = strSafeGetLength(strIn);

        bstrRet = SysAllocStringLen(strIn, uStrSize);

        if (!bstrRet)
        {
            vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
        }
    }

    LOG(LOG_TRACE) << L"strIn " << vbl::Hex((unsigned long)strIn) << L", strOut " << vbl::Hex((unsigned long)strOut)
                   << L", *strOut " << vbl::Hex((unsigned long)*strOut) << L", size " << vbl::Hex((unsigned long)uStrSize)
                   << L", ret " << vbl::Hex((unsigned long)bstrRet);

    if (*strOut)
    {
        SysFreeString(*strOut);
    }

    *strOut = bstrRet;

    LOG(LOG_TRACE) << L"ret = " << vbl::Hex((unsigned long)bstrRet);

    return bstrRet;
} /* __vbaStrCopy */

/**
 * @brief			Converts a Variant (non including errors) to a BSTR.
 * @param			pvargIn			Value to convert.
 * @returns			Result BSTR, nullptr on error (not exactly as VB, but safer).
 */
EXPORT BSTR __stdcall __vbaStrVarCopy(VARIANTARG* pvargIn)
{
    return __vbaStrErrVarCopy(pvargIn);
} /* __vbaStrVarCopy */

/**
 * @brief			Returns pvarSrc's value as a BSTR, borrowing pvarSrc's own bstrVal
 *					when it's already VT_BSTR instead of copying (see the header's
 *					doc comment for the full confirmed behavior/ownership contract,
 *					including the parameter order -- pOutOwned first, pvarSrc second).
 */
EXPORT BSTR __stdcall __vbaStrVarVal(BSTR* pOutOwned, VARIANTARG* pvarSrc)
{

    if (pOutOwned)
    {
        *pOutOwned = nullptr;
    }

    if (!pvarSrc)
    {
        return nullptr;
    }

    LOG(LOG_TRACE) << L"pvarSrc " << vbl::Hex((unsigned long)pvarSrc) << L", pvarSrc->vt " << vbl::Hex((unsigned long)pvarSrc->vt);

    if (pvarSrc->vt == VT_BSTR)
    {
        // Borrowed -- *pOutOwned stays nullptr, signaling the caller must not free
        // this (it belongs to pvarSrc).
        return pvarSrc->bstrVal;
    }

    // Real coercion + a genuine copy -- the caller now owns the result.
    BSTR bstrOwned = __vbaStrVarCopy(pvarSrc);

    if (pOutOwned)
    {
        *pOutOwned = bstrOwned;
    }

    return bstrOwned;
} /* __vbaStrVarVal */

/**
 * @brief			Takes ownership of a Variant's string value, coercing to VT_BSTR
 *					in place first if needed, and empties the Variant to VT_EMPTY
 *					WITHOUT freeing the string -- ownership of the returned BSTR
 *					transfers to the caller.
 * @param			pvarg			Variant to take the string value from. Left as
 *									VT_EMPTY on return.
 * @returns			The BSTR value (now owned by the caller), or nullptr on failure.
 */
EXPORT BSTR __stdcall __vbaStrVarMove(VARIANTARG* pvarg)
{

    if (!pvarg)
    {
        return nullptr;
    }

    LOG(LOG_TRACE) << L"pvarg " << vbl::Hex((unsigned long)pvarg) << L", pvarg->vt " << vbl::Hex((unsigned long)pvarg->vt);

    if (pvarg->vt != VT_BSTR)
    {
        HRESULT hr = VariantChangeType(pvarg, pvarg, 0, VT_BSTR);

        if (FAILED(hr))
        {
            vbaRaiseException(vbaErrorFromHRESULT(hr));
            return nullptr;
        }
    }

    BSTR bstrRet = pvarg->bstrVal;
    pvarg->vt = VT_EMPTY;

    return bstrRet;
} /* __vbaStrVarMove */

/**
 * @brief			Copies a BSTR to a fixed size buffer, right padding with spaces if needed.
 * @param			uiFixedSize		Size of the fixed buffer, in wide characters.
 * @param			lptstrDest		Where the source BSTR value will be copied to. Must be at least uiFixedSize + 1 in size.
 * @param			bstrIn			Source BSTR.
 */
EXPORT void __stdcall __vbaLsetFixstr(UINT uiFixedSize, LPWSTR lptstrDest, BSTR bstrIn)
{
    wchar_t buffer[40] = {0};
    // Create a format string for swprintf to right pad the input string with spaces until the fixed size.
    swprintf(buffer, 39, L"%%-%ds", uiFixedSize);
    // Now do the format
    swprintf(lptstrDest, uiFixedSize + 1, buffer, bstrIn);
} /* __vbaLsetFixstr */

/**
 * @brief			Creates a BSTR from a fixed size buffer.
 * @param			ui				Size of the fixed buffer, in wide characters.
 * @param			strIn			Buffer to copy into the BSTR. Must be at least ui in size.
 * @returns			Result BSTR, exception on error.
 */
EXPORT BSTR __stdcall __vbaStrFixstr(UINT ui, OLECHAR* strIn)
{
    BSTR result = SysAllocStringLen(strIn, ui);

    if (!result)
    {
        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
    }

    return result;
} /* __vbaStrFixstr */

/**
 * @brief			Frees a BSTR via its pointer, and nulls it.
 * @param			pbstrIn			Pointer to a BSTR to free and null it.
 */
EXPORT void __fastcall __vbaFreeStr(BSTR* pbstrIn)
{

    if (pbstrIn)
    {
        // LOG(LOG_TRACE) << L"*pbstrIn " << vbl::Hex(*pbstrIn);

        if (*pbstrIn)
        {
            LOG(LOG_TRACE) << L"*pbstrIn '" << vbl::Bstr(*pbstrIn) << L"'";

            SysFreeString(*pbstrIn);
            *pbstrIn = 0;
        }
    }
} /* __vbaFreeStr */

/**
 * @brief			Frees a list of BSTRs via their pointers, and nulls them.
 * @param			argCount		Count of elements.
 * @param			...				Pointers to BSTRs to free and null them.
 */
EXPORT void __cdecl __vbaFreeStrList(unsigned int argCount, ...)
{
    BSTR* pbstrElement;

    va_list args;
    va_start(args, argCount);

    while (argCount--)
    {
        pbstrElement = va_arg(args, BSTR*);
        LOG(LOG_TRACE) << L"arg() " << vbl::Hex((unsigned long)pbstrElement) << L", argsRemaining " << vbl::Hex((unsigned long)argCount);

        __vbaFreeStr(pbstrElement);
    }

    va_end(args);
} /* __vbaFreeStrList */
