#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

// MIDL-generated header for this object
#include "ErrObject.h"

/*
 * The VB "Err" object (rtcErrObj, ordinal 685). The real DLL returns an AddRef'd pointer to a
 * VBA _ErrObject, which compiled code calls through its vtable (see ErrObject.idl for the layout).
 * The state is simply the last-raised-error globals in Exceptions.cpp
 * (g_iLastErrorCode / g_tLastExcepInfo).
 * TODO: Err should be cleared by Resume / Exit Sub / On Error GoTo 0 (real VB does it); here only
 * Err.Clear does. BSTRs previously stored in g_tLastExcepInfo are leaked on overwrite.
 */
extern vbaExceptions g_iLastErrorCode;
extern EXCEPINFO g_tLastExcepInfo;
const wchar_t* GetVBExceptionDescription(vbaExceptions exception); /* from Exceptions.cpp */

class _ErrObjectImpl final : public _ErrObject
{
    enum : DISPID
    {
        DISP_NUMBER = 0,
        DISP_SOURCE = 0x60030002,
        DISP_DESCRIPTION = 0x60030004,
        DISP_HELPFILE = 0x60030006,
        DISP_HELPCONTEXT = 0x60030008,
        DISP_RAISE = 0x6003000A,
        DISP_CLEAR = 0x6003000B,
        DISP_LASTDLLERROR = 0x6003000C,
    };

    LONG m_refs = 1;

    static BSTR CopyOrEmpty(BSTR s) { return s ? SysAllocStringLen(s, SysStringLen(s)) : SysAllocString(L""); }

    static VARIANT Deref(VARIANT* v)
    {
        VARIANT t = *v;
        while (t.vt == (VT_VARIANT | VT_BYREF) && t.pvarVal)
            t = *t.pvarVal;
        return t;
    }

    /* A missing optional argument (VT_ERROR / VT_EMPTY / VT_NULL / null pointer) means "not supplied". */
    static bool OptStr(VARIANT* v, BSTR* out)
    {
        if (!v)
            return false;
        VARIANT t = Deref(v);
        if (t.vt == VT_ERROR || t.vt == VT_EMPTY || t.vt == VT_NULL)
            return false;
        VARIANT s;
        VariantInit(&s);
        if (FAILED(VariantChangeType(&s, &t, 0, VT_BSTR)))
            return false;
        *out = s.bstrVal;
        return true;
    }

    static bool OptLong(VARIANT* v, long* out)
    {
        if (!v)
            return false;
        VARIANT t = Deref(v);
        if (t.vt == VT_ERROR || t.vt == VT_EMPTY || t.vt == VT_NULL)
            return false;
        VARIANT l;
        VariantInit(&l);
        if (FAILED(VariantChangeType(&l, &t, 0, VT_I4)))
            return false;
        *out = l.lVal;
        return true;
    }

public:
    /* The object is a process-lifetime singleton: reference counting is cosmetic. */
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv)
            return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == IID__ErrObject)
        {
            *ppv = static_cast<_ErrObject*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return (ULONG)InterlockedIncrement(&m_refs); }
    ULONG STDMETHODCALLTYPE Release() override { return (ULONG)InterlockedDecrement(&m_refs); }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override
    {
        if (pctinfo)
            *pctinfo = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo** ppTInfo) override
    {
        if (ppTInfo)
            *ppTInfo = nullptr;
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names, UINT cNames, LCID, DISPID* ids) override
    {
        static const struct
        {
            const wchar_t* n;
            DISPID id;
        } tbl[] = {
            {L"Number", DISP_NUMBER}, {L"Source", DISP_SOURCE}, {L"Description", DISP_DESCRIPTION},
            {L"HelpFile", DISP_HELPFILE}, {L"HelpContext", DISP_HELPCONTEXT}, {L"Raise", DISP_RAISE},
            {L"Clear", DISP_CLEAR}, {L"LastDllError", DISP_LASTDLLERROR},
        };
        HRESULT hr = S_OK;
        for (UINT i = 0; i < cNames; i++)
        {
            ids[i] = DISPID_UNKNOWN;
            if (i > 0)
            {
                hr = DISP_E_UNKNOWNNAME;
                continue;
            }
            bool found = false;
            for (const auto& e : tbl)
            {
                if (_wcsicmp(names[0], e.n) == 0)
                {
                    ids[0] = e.id;
                    found = true;
                }
            }
            if (!found)
                hr = DISP_E_UNKNOWNNAME;
        }
        return hr;
    }
    HRESULT STDMETHODCALLTYPE Invoke(
        DISPID id, REFIID, LCID, WORD flags, DISPPARAMS* dp, VARIANT* res, EXCEPINFO*, UINT*) override
    {
        const bool put = (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0;
        VARIANT arg0;
        VariantInit(&arg0);
        if (put && dp && dp->cArgs >= 1)
            arg0 = dp->rgvarg[0];
        if (res)
            VariantInit(res);

        auto retLong = [&](long v) {
            if (res)
            {
                res->vt = VT_I4;
                res->lVal = v;
            }
            return S_OK;
        };
        auto retStr = [&](BSTR v) {
            if (res)
            {
                res->vt = VT_BSTR;
                res->bstrVal = v;
            }
            else
                SysFreeString(v);
            return S_OK;
        };

        switch (id)
        {
        case DISP_NUMBER:
        case DISP_HELPCONTEXT:
        {
            if (put)
            {
                long v = 0;
                if (!OptLong(&arg0, &v))
                    return DISP_E_TYPEMISMATCH;
                return id == DISP_NUMBER ? put_Number(v) : put_HelpContext(v);
            }
            long v = 0;
            if (id == DISP_NUMBER) get_Number(&v); else get_HelpContext(&v);
            return retLong(v);
        }
        case DISP_LASTDLLERROR:
        {
            long v = 0;
            get_LastDllError(&v);
            return retLong(v);
        }
        case DISP_SOURCE:
        case DISP_DESCRIPTION:
        case DISP_HELPFILE:
        {
            if (put)
            {
                BSTR s = nullptr;
                if (!OptStr(&arg0, &s))
                    return DISP_E_TYPEMISMATCH;
                return id == DISP_SOURCE ? put_Source(s) : id == DISP_DESCRIPTION ? put_Description(s) : put_HelpFile(s);
            }
            BSTR s = nullptr;
            if (id == DISP_SOURCE) get_Source(&s); else if (id == DISP_DESCRIPTION) get_Description(&s); else get_HelpFile(&s);
            return retStr(s);
        }
        case DISP_CLEAR:
            return Clear();
        case DISP_RAISE:
        {
            /* DISPPARAMS arguments are in reverse order: [HelpContext, HelpFile, Description, Source, Number] */
            if (!dp || dp->cArgs < 1 || dp->cArgs > 5)
                return DISP_E_BADPARAMCOUNT;
            VARIANT* a[5] = {};
            for (UINT i = 0; i < dp->cArgs; i++)
                a[dp->cArgs - 1 - i] = &dp->rgvarg[i];
            long number = 0;
            if (!OptLong(a[0], &number))
                return DISP_E_TYPEMISMATCH;
            return Raise(number, a[1], a[2], a[3], a[4]);
        }
        }
        return DISP_E_MEMBERNOTFOUND;
    }

    HRESULT STDMETHODCALLTYPE get_Number(long* rhs) override
    {
        *rhs = (long)g_iLastErrorCode;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE put_Number(long rhs) override
    {
        g_iLastErrorCode = (vbaExceptions)rhs;
        g_tLastExcepInfo.scode = rhs;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE get_Source(BSTR* rhs) override
    {
        *rhs = CopyOrEmpty(g_tLastExcepInfo.bstrSource);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE put_Source(BSTR rhs) override
    {
        g_tLastExcepInfo.bstrSource = rhs;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE get_Description(BSTR* rhs) override
    {
        *rhs = CopyOrEmpty(g_tLastExcepInfo.bstrDescription);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE put_Description(BSTR rhs) override
    {
        g_tLastExcepInfo.bstrDescription = rhs;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE get_HelpFile(BSTR* rhs) override
    {
        *rhs = CopyOrEmpty(g_tLastExcepInfo.bstrHelpFile);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE put_HelpFile(BSTR rhs) override
    {
        g_tLastExcepInfo.bstrHelpFile = rhs;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE get_HelpContext(long* rhs) override
    {
        *rhs = (long)g_tLastExcepInfo.dwHelpContext;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE put_HelpContext(long rhs) override
    {
        g_tLastExcepInfo.dwHelpContext = (DWORD)rhs;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Raise(long Number, VARIANT* Source, VARIANT* Description, VARIANT* HelpFile, VARIANT* HelpContext) override
    {
        EXCEPINFO ei = {0};
        ei.scode = Number;
        long ctx = 0;
        OptStr(Source, &ei.bstrSource);
        OptStr(HelpFile, &ei.bstrHelpFile);
        if (OptLong(HelpContext, &ctx))
            ei.dwHelpContext = (DWORD)ctx;
        if (!OptStr(Description, &ei.bstrDescription))
        {
            const wchar_t* d = GetVBExceptionDescription((vbaExceptions)Number);
            ei.bstrDescription = SysAllocString(d ? d : L"Application-defined or object-defined error");
        }
        vbaRaiseException((vbaExceptions)Number, &ei); /* does not return */
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Clear() override
    {
        g_iLastErrorCode = VBA_EXCEPTION_NO_ERROR;
        memset(&g_tLastExcepInfo, 0, sizeof(g_tLastExcepInfo));
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE get_LastDllError(long* rhs) override
    {
        *rhs = (long)GetLastError();
        return S_OK;
    }
};

/**
 * @brief			Returns the Err object, AddRef'd. Confirmed via real msvbvm60.dll disassembly:
 *					gets the per-thread object, AddRefs it and returns it.
 */
EXPORT _ErrObject* __stdcall rtcErrObj()
{
    static _ErrObjectImpl s_err;
    s_err.AddRef();
    return &s_err;
} /* rtcErrObj */
