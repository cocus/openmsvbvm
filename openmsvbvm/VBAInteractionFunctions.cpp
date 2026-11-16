#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include "StringManipulation.hpp"

#include <strsafe.h>
#include <string>
#include <cwctype>

/**
 * @brief			Sets the focus to the specified window with the corresponding title.
 * @param			pvargTitle		Title of the window.
 * @param			pvargWait		Bool that specifies to wait for user input or not.
 */
EXPORT void __stdcall rtcAppActivate(VARIANTARG* pvargTitle, VARIANTARG* pvargWait)
{

    LOG(LOG_DEBUG) << L"pvargTitle " << vbl::Hex((unsigned long)pvargTitle) << L", pvargWait " << vbl::Hex((unsigned long)pvargWait);

    UINT uiWaitTime = 0;
    BSTR strTitle;

    /* Create a copy of pvargTitle and convert it to BSTR */
    strTitle = __vbaStrErrVarCopy(pvargTitle);

    LOG(LOG_DEBUG) << L"Window title: '" << vbl::Bstr(strTitle) << L"', wait time " << vbl::Hex(uiWaitTime);

    HWND hWnd = FindWindowW(nullptr, strTitle);

    SysFreeString(strTitle);

    if (!hWnd)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return;
    }

    LOG(LOG_DEBUG) << L"Window hWnd " << vbl::Hex((unsigned long)hWnd);

    DWORD dwProcessId;
    DWORD dwThreadProcessId;

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
EXPORT void __stdcall rtcChoose(VARIANTARG* pvargDest, FLOAT fIndex, SAFEARRAY** ppsaArray)
{
    VARIANTARG vargLocalCopy;
    VARIANTARG pv;
    LONG lIndices = 0;
    HRESULT hr;

    /* Zero out the destination variant variable */
    memset(pvargDest, 0, sizeof(VARIANTARG));

    pvargDest->vt = VT_NULL;

    /* Check for index bounds */
    if (!(fIndex >= 0.0 && fIndex < 2147483600.0))
    {
        // TODO: Raise an exception?
        vbaRaiseException(VBA_EXCEPTION_INVALID_USE_OF_NULL);
        return;
    }

    /* If the array is more than one dimension, raise an exception */
    if (SafeArrayGetDim(*ppsaArray) != 1)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
        return;
    }

#pragma warning(disable : 244) /* We know that we'll loose some data by converting, don't yell at me because of that */
    /* For some weird reason VB passes a float instead of a LONG, so convert that to LONG */
    lIndices = fIndex - 1;
#pragma warning(default : 244) /* Restore previous warning */

    /* Get the element from the array */
    hr = SafeArrayGetElement(*ppsaArray, &lIndices, (void*)&pv);

    if (hr < 0)
    {
        vbaRaiseException(vbaErrorFromHRESULT(hr));
        return;
    }

    if (!(pv.vt & VT_BYREF))
    {
        /* VT_BYREF not set */
        memcpy(pvargDest, &pv, sizeof(VARIANTARG));
    }
    else
    {
        /* VT_BYREF set */
        hr = VariantCopyInd(&vargLocalCopy, &pv);

        if (hr < 0)
        {
            vbaRaiseException(vbaErrorFromHRESULT(hr));
            return;
        }
        else
        {
            memcpy(pvargDest, &vargLocalCopy, sizeof(VARIANTARG));
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
EXPORT BSTR __stdcall rtcEnvironBstr(VARIANTARG* pvargIn)
{

    BSTR bstrRet;
    BSTR bstrEnvName;

    bstrEnvName = __vbaStrVarCopy(pvargIn);

    if (!bstrEnvName)
    {
        bstrRet = SysAllocString(L"");
    }
    else
    {
        LOG(LOG_DEBUG) << L"EnvironmentVariable: '" << vbl::Bstr(bstrEnvName) << L"'";

        DWORD dwSize;
        dwSize = GetEnvironmentVariableW(bstrEnvName, NULL, 0);

        if (dwSize == 0)
        {
            /* Specified environment string not found */
            bstrRet = SysAllocString(L"");
        }
        else
        {
            wchar_t* buffer = new wchar_t[dwSize];

            if (!buffer)
            {
                vbaRaiseException(VBA_EXCEPTION_OUT_OF_MEMORY);
                return NULL;
            }

            dwSize = GetEnvironmentVariableW(bstrEnvName, buffer, dwSize);

            bstrRet = SysAllocStringLen(buffer, dwSize);

            delete[] buffer;

            if (!bstrRet)
            {
                vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
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
EXPORT VARIANTARG* __stdcall rtcEnvironVar(VARIANTARG* pvargOut, VARIANTARG* pvargIn)
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
EXPORT VARIANTARG* __stdcall rtcImmediateIf(VARIANTARG* pvargDest, VARIANTARG* pvargBooleanValue, VARIANTARG* pvargTrueValue, VARIANTARG* pvargFalseValue)
{
    HRESULT hr;
    VARIANTARG vargB;

    /* Convert pvargBooleanValue to a Boolean variant (deref if necessary) */
    hr = VariantChangeType(&vargB, pvargBooleanValue, 0, VT_BOOL);

    if (hr != S_OK)
    {
        vbaRaiseException(vbaErrorFromHRESULT(hr));
        return pvargDest;
    }

    LOG(LOG_DEBUG) << L"booleanValue " << vargB.boolVal << L", trueValueVT " << vbl::Hex(pvargTrueValue->vt) << L", falseValueVT "
                   << vbl::Hex(pvargFalseValue->vt);

    /* Create a copy of pvargTrueValue or pvargFalseValue as needed */
    hr = VariantCopyInd(pvargDest, vargB.boolVal ? pvargTrueValue : pvargFalseValue);

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
 * @param			pvargStart		Whole number that is the start of the overall range of numbers. The number can't be less than  0.
 * @param			pvargEnd		Whole number that is the end of the overall range of numbers. The number can't be equal to or  less than start.
 * @param			pvargInterval	Whole number that specifies the interval.
 */
EXPORT void __stdcall rtcPartition(VARIANTARG* pvargDest, VARIANTARG* pvargNumber, VARIANTARG* pvargStart, VARIANTARG* pvargEnd, VARIANTARG* pvargInterval)
{

    VARIANTARG vargNumber;
    VARIANTARG vargStart;
    VARIANTARG vargEnd;
    VARIANTARG vargInterval;

    int iRange1, iRange2;

    HRESULT hr;

    VariantInit(&vargNumber);
    VariantInit(&vargStart);
    VariantInit(&vargEnd);
    VariantInit(&vargInterval);

    /* Convert Number to I4 */
    hr = VariantChangeType(&vargNumber, pvargNumber, 0, VT_I4);

    if (hr != S_OK)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return;
    }

    /* Convert Start to I4 */
    hr = VariantChangeType(&vargStart, pvargStart, 0, VT_I4);

    if (hr != S_OK)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return;
    }

    /* Convert End to I4 */
    hr = VariantChangeType(&vargEnd, pvargEnd, 0, VT_I4);

    if (hr != S_OK)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return;
    }

    /* Convert Interval to I4 */
    hr = VariantChangeType(&vargInterval, pvargInterval, 0, VT_I4);

    if (hr != S_OK)
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
        return;
    }

    LOG(LOG_DEBUG) << L"Number " << vargNumber.intVal << L", Start " << vargStart.intVal << L", End " << vargEnd.intVal
                   << L", Interval " << vargInterval.intVal;

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

/* ---- SendKeys (port of Kr00l's https://github.com/cocus/openmsvbvm/issues/6#issuecomment-6923532826) ----
 * Compile-time settings (define to 0 or 1, e.g. in project preprocessor definitions):
 *   OPENMSVBVM_SENDKEYS_UNICODE    1 (default) = Unicode chars via KEYEVENTF_UNICODE,
 *                                  0 = legacy byte-based VBA.SendKeys compatibility
 *   OPENMSVBVM_SENDKEYS_EXTENSIONS 1 (default) = extra key names ({SLEEP n}, {DOEVENTS}, numpad, F17-F24...),
 *                                  0 = strict VBA.SendKeys syntax
 */
#ifndef OPENMSVBVM_SENDKEYS_UNICODE
#define OPENMSVBVM_SENDKEYS_UNICODE 1
#endif
#ifndef OPENMSVBVM_SENDKEYS_EXTENSIONS
#define OPENMSVBVM_SENDKEYS_EXTENSIONS 1
#endif

namespace
{
struct SkMods
{
    bool shift, ctrl, alt;
};

static int g_skError = 0; /* first error raised while sending, 0 = none */

static void SkKey(WORD vk, WORD scan, DWORD flags)
{
    if (g_skError)
        return;
    INPUT in = {};
    in.type = INPUT_KEYBOARD;
    in.ki.wVk = vk;
    in.ki.wScan = scan;
    in.ki.dwFlags = flags;
    SetLastError(0);
    if (SendInput(1, &in, sizeof(in)) != 1)
        g_skError = GetLastError() != 0 ? VBA_EXCEPTION_INVALID_PROCEDURE_CALL : VBA_EXCEPTION_PERMISSION_DENIED;
}

static void SkPump()
{
    MSG msg;
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

static void SkChar(wchar_t ch, SkMods m, long repeat = 1)
{
    SHORT vk;
#if OPENMSVBVM_SENDKEYS_UNICODE
    vk = VkKeyScanW(ch);
    if (vk == -1)
    {
        for (long i = 0; i < repeat; i++)
        {
            SkKey(0, ch, KEYEVENTF_UNICODE);
            SkKey(0, ch, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP);
        }
        return;
    }
#else
    char a[4] = {};
    WideCharToMultiByte(CP_ACP, 0, &ch, 1, a, sizeof(a), nullptr, nullptr);
    vk = VkKeyScanW((unsigned char)a[0]);
    if (vk == -1)
        vk = (unsigned char)a[0]; /* byte-as-VK fallback */
#endif
    int mods = (vk & 0xFF00) >> 8;
    WORD key = (WORD)(vk & 0xFF);
    if (!m.shift && (mods & 1))
        SkKey(VK_SHIFT, 0, 0);
    if (!m.ctrl && (mods & 2))
        SkKey(VK_CONTROL, 0, 0);
    if (!m.alt && (mods & 4))
        SkKey(VK_MENU, 0, 0);
    for (long i = 0; i < repeat; i++)
    {
        SkKey(key, 0, 0);
        SkKey(key, 0, KEYEVENTF_KEYUP);
    }
    if (!m.shift && (mods & 1))
        SkKey(VK_SHIFT, 0, KEYEVENTF_KEYUP);
    if (!m.ctrl && (mods & 2))
        SkKey(VK_CONTROL, 0, KEYEVENTF_KEYUP);
    if (!m.alt && (mods & 4))
        SkKey(VK_MENU, 0, KEYEVENTF_KEYUP);
}

static WORD SkNamedKey(const std::wstring& name, bool& ext)
{
    static const struct
    {
        const wchar_t* n;
        WORD vk;
        bool ext;
        bool extension;
    } tbl[] = {
        {L"ADD", VK_ADD, false, true},
        {L"NUMPADADD", VK_ADD, false, true},
        {L"BS", VK_BACK, false, false},
        {L"BKSP", VK_BACK, false, false},
        {L"BACKSPACE", VK_BACK, false, false},
        {L"BREAK", VK_CANCEL, true, false},
        {L"CAPSLOCK", VK_CAPITAL, false, false},
        {L"CLEAR", VK_CLEAR, false, false},
        {L"DECIMAL", VK_DECIMAL, false, true},
        {L"NUMPADDECIMAL", VK_DECIMAL, false, true},
        {L"DEL", VK_DELETE, true, false},
        {L"DELETE", VK_DELETE, true, false},
        {L"DIVIDE", VK_DIVIDE, true, true},
        {L"NUMPADDIVIDE", VK_DIVIDE, true, true},
        {L"DOWN", VK_DOWN, true, false},
        {L"END", VK_END, true, false},
        {L"ENTER", VK_RETURN, false, false},
        {L"RETURN", VK_RETURN, false, true},
        {L"ESC", VK_ESCAPE, false, false},
        {L"ESCAPE", VK_ESCAPE, false, false},
        {L"HELP", VK_HELP, false, false},
        {L"HOME", VK_HOME, true, false},
        {L"INS", VK_INSERT, true, false},
        {L"INSERT", VK_INSERT, true, false},
        {L"LEFT", VK_LEFT, true, false},
        {L"MULTIPLY", VK_MULTIPLY, false, true},
        {L"NUMPADMULTIPLY", VK_MULTIPLY, false, true},
        {L"NUMLOCK", VK_NUMLOCK, true, false},
        {L"NUMPAD0", VK_NUMPAD0, false, true},
        {L"NUMPAD1", VK_NUMPAD1, false, true},
        {L"NUMPAD2", VK_NUMPAD2, false, true},
        {L"NUMPAD3", VK_NUMPAD3, false, true},
        {L"NUMPAD4", VK_NUMPAD4, false, true},
        {L"NUMPAD5", VK_NUMPAD5, false, true},
        {L"NUMPAD6", VK_NUMPAD6, false, true},
        {L"NUMPAD7", VK_NUMPAD7, false, true},
        {L"NUMPAD8", VK_NUMPAD8, false, true},
        {L"NUMPAD9", VK_NUMPAD9, false, true},
        {L"NUMPADENTER", VK_RETURN, true, true},
        {L"KEYPADENTER", VK_RETURN, true, true},
        {L"KPENTER", VK_RETURN, true, true},
        {L"PGDN", VK_NEXT, true, false},
        {L"PGUP", VK_PRIOR, true, false},
        {L"PAUSE", VK_PAUSE, false, true},
        {L"PRINT", VK_PRINT, true, true},
        {L"PRTSC", VK_SNAPSHOT, true, false},
        {L"PRINTSCREEN", VK_SNAPSHOT, true, true},
        {L"RIGHT", VK_RIGHT, true, false},
        {L"SCROLLLOCK", VK_SCROLL, false, false},
        {L"SELECT", VK_SELECT, false, true},
        {L"SEPARATOR", VK_SEPARATOR, false, true},
        {L"NUMPADSEPARATOR", VK_SEPARATOR, false, true},
        {L"SPACE", VK_SPACE, false, true},
        {L"SUBTRACT", VK_SUBTRACT, false, true},
        {L"NUMPADSUBTRACT", VK_SUBTRACT, false, true},
        {L"TAB", VK_TAB, false, false},
        {L"UP", VK_UP, true, false},
    };
    std::wstring up = name;
    for (auto& c : up)
        c = (wchar_t)towupper(c);
    for (const auto& e : tbl)
    {
        if (up == e.n)
        {
            if (e.extension && !OPENMSVBVM_SENDKEYS_EXTENSIONS)
                return 0;
            ext = e.ext;
            return e.vk;
        }
    }
    /* F1..F16 (F17..F24 are extensions) */
    if (up.size() >= 2 && up[0] == L'F' && iswdigit(up[1]))
    {
        int n = _wtoi(up.c_str() + 1);
        bool digitsOnly = up.find_first_not_of(L"0123456789", 1) == std::wstring::npos;
        if (digitsOnly && n >= 1 && (n <= 16 || (OPENMSVBVM_SENDKEYS_EXTENSIONS && n <= 24)))
        {
            ext = false;
            return (WORD)(VK_F1 + n - 1);
        }
    }
    return 0;
}

static bool SkSequence(const std::wstring& t, size_t& pos, bool stopAtClose, SkMods m);

static bool SkNamed(const std::wstring& t, size_t& pos, SkMods m)
{
    size_t close = t.find(L'}', pos + 2); /* allow {}} */
    if (close == std::wstring::npos)
        return false;
    std::wstring token = t.substr(pos + 1, close - pos - 1);
    if (token.empty())
        return false;
    if (token.size() == 1)
    {
        SkChar(token[0], m);
        pos = close;
        return true;
    }
    for (size_t p; (p = token.find(L"  ", 1)) != std::wstring::npos;)
        token.erase(p, 1);
    size_t rp = token.find(L' ', 1);
    std::wstring keyName;
    if (rp != std::wstring::npos)
    {
        if (token.find(L' ', rp + 1) != std::wstring::npos)
            return false;
        keyName = token.substr(0, rp);
    }
    else
        keyName = token;

    WORD key = 0;
    bool ext = false;
    long repeat = 1;
    bool isSleep = false, isDoEvents = false;
#if OPENMSVBVM_SENDKEYS_EXTENSIONS
    std::wstring up = keyName;
    for (auto& c : up)
        c = (wchar_t)towupper(c);
    if (up == L"SLEEP" || up == L"DELAY")
    {
        if (rp == std::wstring::npos)
            return false;
        isSleep = true;
    }
    else if (up == L"DOEVENTS" || up == L"NOOP")
    {
        if (rp != std::wstring::npos)
            return false;
        if (up == L"NOOP")
        {
            pos = close;
            return true;
        }
        isDoEvents = true;
    }
#endif
    if (!isSleep && !isDoEvents)
    {
        key = SkNamedKey(keyName, ext);
        if (key == 0 && keyName.size() != 1)
            return false;
    }
    if (rp != std::wstring::npos)
    {
        std::wstring rep = token.substr(rp + 1);
        if (rep.empty())
            return false;
        size_t i = rep.find_first_not_of(L'0');
        if (i == std::wstring::npos)
            repeat = 0;
        else
        {
            rep = rep.substr(i);
            if (rep.size() > 5 || rep.find_first_not_of(L"0123456789") != std::wstring::npos)
                return false;
            repeat = _wtol(rep.c_str());
            if (repeat > 10000)
                return false;
        }
    }
    if (key == 0)
    {
        if (isSleep)
        {
            if (repeat > 0)
                Sleep((DWORD)repeat);
        }
        else if (isDoEvents)
            SkPump();
        else
            SkChar(keyName[0], m, repeat);
    }
    else
    {
        for (long i = 0; i < repeat; i++)
        {
            DWORD f = ext ? KEYEVENTF_EXTENDEDKEY : 0;
            SkKey(key, 0, f);
            SkKey(key, 0, f | KEYEVENTF_KEYUP);
        }
    }
    pos = close;
    return true;
}

static bool SkModified(const std::wstring& t, size_t& pos, SkMods m)
{
    bool shift = false, ctrl = false, alt = false, result = true, cancel = false;
    while (pos < t.size())
    {
        wchar_t c = t[pos];
        if (c == L'+')
        {
            if (m.shift || shift)
            {
                cancel = true;
                result = false;
                break;
            }
            shift = true;
            SkKey(VK_SHIFT, 0, 0);
        }
        else if (c == L'^')
        {
            if (m.ctrl || ctrl)
            {
                cancel = true;
                result = false;
                break;
            }
            ctrl = true;
            SkKey(VK_CONTROL, 0, 0);
        }
        else if (c == L'%')
        {
            if (m.alt || alt)
            {
                cancel = true;
                result = false;
                break;
            }
            alt = true;
            SkKey(VK_MENU, 0, 0);
        }
        else
            break;
        pos++;
    }
    if (!cancel)
    {
        SkMods n = {m.shift || shift, m.ctrl || ctrl, m.alt || alt};
        if (pos < t.size())
        {
            switch (t[pos])
            {
            case L'{':
                result = SkNamed(t, pos, n);
                break;
            case L'}':
            case L')':
                result = false;
                break;
            case L'(':
                pos++;
                result = SkSequence(t, pos, true, n);
                break;
            case L'~':
            {
                size_t p = 0;
                result = SkNamed(L"{ENTER}", p, SkMods{});
                break;
            }
            default:
                SkChar(t[pos], n);
                result = true;
                break;
            }
        }
    }
    if (shift && GetAsyncKeyState(VK_SHIFT) < 0)
        SkKey(VK_SHIFT, 0, KEYEVENTF_KEYUP);
    if (ctrl && GetAsyncKeyState(VK_CONTROL) < 0)
        SkKey(VK_CONTROL, 0, KEYEVENTF_KEYUP);
    if (alt && GetAsyncKeyState(VK_MENU) < 0)
        SkKey(VK_MENU, 0, KEYEVENTF_KEYUP);
    return result;
}

static bool SkSequence(const std::wstring& t, size_t& pos, bool stopAtClose, SkMods m)
{
    while (pos < t.size())
    {
        wchar_t c = t[pos];
        switch (c)
        {
        case L'+':
        case L'^':
        case L'%':
            if (!SkModified(t, pos, m))
                return false;
            break;
        case L'{':
            if (!SkNamed(t, pos, m))
                return false;
            break;
        case L'}':
            return false;
        case L'(':
            pos++;
            if (!SkSequence(t, pos, true, m))
                return false;
            if (pos >= t.size() || t[pos] != L')')
                return false;
            break;
        case L')':
            return stopAtClose;
        case L'~':
        {
            size_t p = 0;
            if (!SkNamed(L"{ENTER}", p, SkMods{}))
                return false;
            break;
        }
        default:
            SkChar(c, m);
        }
        pos++;
    }
    return !stopAtClose;
}
} // namespace

/**
 * @brief			Sends multiple keys or combination of keys to the focused window.
 * @param			bstrKeys		BSTR with the keys and combination of keys.
 * @param			pvargWaitTime	Optional Wait flag; retained for compatibility only (SendInput can't wait).
 */
EXPORT void __stdcall rtcSendKeys(BSTR bstrKeys, VARIANTARG* pvargWaitTime)
{
    LOG(LOG_DEBUG) << L"inputKeys '" << vbl::Bstr(bstrKeys) << L"'";

    std::wstring text = bstrKeys ? std::wstring(bstrKeys, SysStringLen(bstrKeys)) : std::wstring();
    g_skError = 0;

    /* Temporarily turn off CapsLock so shifted characters come out right */
    bool caps = (GetKeyState(VK_CAPITAL) & 1) != 0;
    if (caps)
    {
        SkKey(VK_CAPITAL, 0, 0);
        SkKey(VK_CAPITAL, 0, KEYEVENTF_KEYUP);
    }
    size_t pos = 0;
    bool ok = SkSequence(text, pos, false, SkMods{});
    if (caps)
    {
        int saved = g_skError;
        g_skError = 0;
        SkKey(VK_CAPITAL, 0, 0);
        SkKey(VK_CAPITAL, 0, KEYEVENTF_KEYUP);
        if (!saved)
            saved = g_skError;
        g_skError = saved;
    }
    int err = g_skError;
    g_skError = 0;
    if (err)
        vbaRaiseException(static_cast<vbaExceptions>(err));
    else if (!ok)
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
} /* rtcSendKeys */

/**
 * @brief			Implements Timer: seconds elapsed since local midnight, with milliseconds.
 * @note			Confirmed via real msvbvm60.dll disassembly (ordinal 535): GetLocalTime, then
 *					(hour*60+minute)*60+second plus milliseconds*0.001, returned in ST(0).
 *					(dumpbin's listing shows the final operation as an fsubp; read as a
 *					subtraction that would make Timer go backwards, so this assumes the
 *					operand-order quirk of fsubp/fsubrp disassembly and adds.)
 * @returns			Seconds since midnight as a double (the caller rounds it to Single).
 */
EXPORT double __stdcall rtcGetTimer()
{
    SYSTEMTIME st;
    GetLocalTime(&st);
    double dResult = (double)(((st.wHour * 60 + st.wMinute) * 60) + st.wSecond) + st.wMilliseconds * 0.001;
    /* The compiled VB code tests the FPU status word right after Timer's __vbaFpR4 and raises
       "Expression too complex"/__vbaFPException on any sticky exception flag, including stale
       ones left behind by earlier FPU work, so clear them before handing the value back. */
    unsigned int uiBefore = _clearfp();
    LOG(LOG_DEBUG) << L"rtcGetTimer: seconds*1000 = " << (long)(dResult * 1000.0) << L", stale FPU status flags cleared = "
                   << vbl::Hex(uiBefore);
    return dResult;
} /* rtcGetTimer */

/**
 * @brief			Implements DoEvents(): processes pending window messages.
 * @note			ALMOST A STUB: it only pumps the thread's message queue. The real DLL does
 *					more (it goes through its own internal message-dispatch helpers) and
 *					returns the number of open forms; this always returns 0.
 * @returns			Always 0 (TODO: number of open forms).
 */
EXPORT short __stdcall rtcDoEvents()
{
    MSG msg;
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
    {
        if (msg.message == WM_QUIT)
        {
            PostQuitMessage((int)msg.wParam); /* keep the quit request for the outer loop */
            break;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
} /* rtcDoEvents */

/**
 * @brief			Starts up a new process.
 * @param			vargPath		Path to the application.
 * @param			uiShowCmd		Show mode for the newly created app.
 * @return			PID of the newly created process on success, -1 otherwise.
 */
EXPORT DWORD __stdcall rtcShell(VARIANTARG* vargPath, unsigned int uiShowCmd)
{
    BSTR bstrCmdLine;
    BOOL processCreated;
    STARTUPINFOW StartupInfo;
    PROCESS_INFORMATION ProcessInformation;

    DWORD lastError;

    LOG(LOG_DEBUG) << L"vargPath " << vbl::Hex((unsigned long)vargPath) << L", uiShowCmd " << vbl::Hex(uiShowCmd);

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
    processCreated = CreateProcessW(0, bstrCmdLine, 0, 0, 0, 0, 0, 0, &StartupInfo, &ProcessInformation);

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
