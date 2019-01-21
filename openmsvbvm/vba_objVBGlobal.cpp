#include "vba_objVBGlobal.h"

#include "vba_internal.h"
#include "vba_exception.h"

HRESULT _stdcall CVBGlobal::Load(IDispatch * object)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("object %.8x", (unsigned int)object);
	return E_NOTIMPL;
}

HRESULT _stdcall CVBGlobal::Unload(IDispatch * object)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("object %.8x", (unsigned int)object);

	return E_NOTIMPL;
}

HRESULT _stdcall CVBGlobal::get_App(IApp ** pdispRetVal)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("pdispRetVal %.8x", (unsigned int)pdispRetVal);

	*pdispRetVal = new CApp();
	return S_OK;
}


HRESULT __stdcall CVBGlobal::QueryInterface(
	REFIID riid,
	void **ppObj
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ(
		"ppObj %.8x", 
		(unsigned int)ppObj
	);
	
	DEBUG_WIDE_GUID(riid, "riid");

	if (riid == IID_IUnknown)
	{
		*ppObj = static_cast<void*>(this);
		AddRef();
		return S_OK;
	}
	if (riid == IID_IVBGlobal)
	{
		*ppObj = static_cast<void*>(this);
		AddRef();
		return S_OK;
	}

	*ppObj = NULL;
	return E_NOINTERFACE;
}

ULONG __stdcall CVBGlobal::AddRef()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall CVBGlobal::Release()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	long nRefCount = 0;
	nRefCount = InterlockedDecrement(&m_nRefCount);
	if (nRefCount == 0) delete this;
	return nRefCount;
}

extern  ULONG g_Components;		/* from DllObjectInterface.cpp */

CVBGlobal::CVBGlobal() : m_nRefCount(1)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	InterlockedIncrement((LONG*)&g_Components);
}

CVBGlobal::~CVBGlobal()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	InterlockedDecrement((LONG*)&g_Components);
}