#include "vba_internal.h"
#include "Logging.hpp"
#include "ControlObject.hpp"

#include <cstring>

// MIDL-generated header for the real _Label interface (Label.idl) -- included
// plainly, the same way App.cpp/FormWrapper.cpp include their own generated
// headers, so _Label comes in as a genuine C++ abstract class.
#include "Label.h"

/**
 * Unlike _FormImpl (FormWrapper.cpp), this is a real, standalone COM object with its
 * own genuine identity -- it's not layered onto an existing raw vtable blob, so no
 * "this is secretly something else" trick is needed here; a normal AddRef/Release
 * refcount and a normal HWND member are enough. See ControlObject.hpp's header
 * comment for the rest of the design.
 */
class _LabelImpl : public _Label
{
public:
    _LabelImpl(HWND hwndControl, const char* pszName) : m_nRefCount(1), m_hwnd(hwndControl)
    {
        strncpy_s(m_name, sizeof(m_name), pszName ? pszName : "", _TRUNCATE);
    }

    // IUnknown
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (!ppvObject)
        {
            return E_POINTER;
        }

        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
        {
            *ppvObject = static_cast<_Label*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG __stdcall AddRef() override
    {
        return InterlockedIncrement(&m_nRefCount);
    }

    ULONG __stdcall Release() override
    {
        long nRefCount = InterlockedDecrement(&m_nRefCount);
        if (nRefCount == 0)
        {
            delete this;
        }
        return nRefCount;
    }

    // IDispatch -- not needed for the direct vtable-slot property access this
    // project exercises today; stubbed like the rest of this class.
    HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo) override
    {
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo) override
    {
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid) override
    {
        return E_NOTIMPL;
    }

    HRESULT __stdcall Invoke(
        DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override
    {
        return E_NOTIMPL;
    }

    HRESULT __stdcall Missing1(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing2(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing3(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing4(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing5(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing6(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing7(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing8(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing9(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing10(void) override
    {
        return E_NOTIMPL;
    }
    HRESULT __stdcall Missing11(void) override
    {
        return E_NOTIMPL;
    }

    // MIDL always names a [propget]/[propput] member's C++ override get_X/put_X,
    // regardless of the plain name ("Name", "Caption") the .idl declares it under
    // (confirmed via Form.h's own generated overrides).
    HRESULT __stdcall get_Name(BSTR* rhs) override
    {
        if (!rhs)
        {
            return E_POINTER;
        }

        wchar_t wszName[128];
        MultiByteToWideChar(CP_ACP, 0, m_name, -1, wszName, 128);

        *rhs = SysAllocString(wszName);
        return *rhs ? S_OK : E_OUTOFMEMORY;
    }

    HRESULT __stdcall Missing12(void) override
    {
        return E_NOTIMPL;
    }

    HRESULT __stdcall get_Caption(BSTR* rhs) override
    {
        if (!rhs)
        {
            return E_POINTER;
        }

        wchar_t caption[512] = {0};
        if (m_hwnd)
        {
            GetWindowTextW(m_hwnd, caption, ARRAYSIZE(caption));
        }

        *rhs = SysAllocString(caption);
        return *rhs ? S_OK : E_OUTOFMEMORY;
    }

    HRESULT __stdcall put_Caption(BSTR rhs) override
    {
        if (m_hwnd)
        {
            SetWindowTextW(m_hwnd, rhs ? rhs : L"");
        }

        return S_OK;
    }

private:
    long m_nRefCount;
    HWND m_hwnd;
    char m_name[64];
};

IDispatch* CreateControlAccessorObject(HWND hwndControl, const char* pszControlName)
{
    _LabelImpl* pImpl = new _LabelImpl(hwndControl, pszControlName);
    return static_cast<IDispatch*>(pImpl);
}
