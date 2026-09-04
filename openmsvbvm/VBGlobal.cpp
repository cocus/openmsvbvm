#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"
#include "ObjectManipulation.hpp"

// MIDL-generated header for this object
#include "VBGlobal.h"

extern ULONG g_Components; /* from DllObjectInterface.cpp */

class VBGlobalImpl : public VBGlobal
{
    LONG refCount = 1;

public:
    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
    {
        if (!ppvObject)
            return E_POINTER;
        *ppvObject = nullptr;

        if (riid == IID_IUnknown || riid == IID_VBGlobal)
        {
            *ppvObject = static_cast<VBGlobal*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        InterlockedIncrement(&g_Components);
        return InterlockedIncrement(&refCount);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG r = InterlockedDecrement(&refCount);
        InterlockedDecrement(&g_Components);
        if (r == 0)
        {
            delete this;
        }
        return r;
    }

    /* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE Load(
        /* [in] */ IDispatch* object)
    {

        LOG_OBJ(LOG_TRACE, this) << L"object = " << vbl::Hex(object);

        return VBFormLoad(object); // TODO: Form or Control
    }

    /* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE Unload(
        /* [in] */ IDispatch* object)
    {

        LOG_OBJ(LOG_TRACE, this) << L"object = " << vbl::Hex(object);

        return VBFormUnload(object); // TODO: Form or Control
    }

    /* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_App(
        /* [retval][out] */ _App** pdispRetVal)
    {

        LOG_OBJ(LOG_TRACE, this) << L"pdispRetVal = " << vbl::Hex(pdispRetVal);

        extern HRESULT Create_App_Instance(IUnknown * *ppApp);

        HRESULT hr = Create_App_Instance(reinterpret_cast<IUnknown**>(pdispRetVal));

        return hr;
    }
};

HRESULT Create_VBGlobal_Instance(IUnknown** ppApp)
{
    if (!ppApp)
        return E_POINTER;
    *ppApp = new VBGlobalImpl();
    if (!*ppApp)
        return E_OUTOFMEMORY;
    (*ppApp)->AddRef();
    return S_OK;
}
