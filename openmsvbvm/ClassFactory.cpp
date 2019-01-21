#include "ClassFactory.h"

#include "vba_objApp.h"
#include "vba_objVBGlobal.h"



extern ULONG g_ServerLocks;

CMasterFactory::CMasterFactory() : m_nRefCount(1)
{
}

CMasterFactory::~CMasterFactory()
{
}

HRESULT __stdcall CMasterFactory::QueryInterface(REFIID riid, void ** ppObj)
{
	HRESULT rc = S_OK;

	if (riid == IID_IUnknown)
	{
		*ppObj = (IUnknown*)this;
	}
	else if (riid == IID_IClassFactory)
	{
		*ppObj = (IClassFactory*)this;
	}
	/*else if (riid == IID_IVBGlobal)
	{
		
	}*/
	else rc = E_NOINTERFACE;

	if (rc == S_OK)
		this->AddRef();

	return rc;
}

ULONG __stdcall CMasterFactory::AddRef()
{
	InterlockedIncrement(&m_nRefCount);

	return this->m_nRefCount;
}

ULONG __stdcall CMasterFactory::Release()
{
	InterlockedDecrement(&m_nRefCount);

	if (this->m_nRefCount == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return this->m_nRefCount;
	}
}

HRESULT __stdcall CMasterFactory::CreateInstance(
	IUnknown * pUnknownOuter,
	const IID & iid,
	void ** ppv)
{
	HRESULT rc = E_UNEXPECTED;

	if (pUnknownOuter != NULL)
	{
		rc = CLASS_E_NOAGGREGATION;
	}
	else if (iid == IID_IVBGlobal)
	{
		IUnknown* p;
		if ((p = new CVBGlobal()) == NULL)
		{
			rc = E_OUTOFMEMORY;
		}
		else {
			rc = p->QueryInterface(iid, ppv);
			p->Release();
		}
	}
	else if (iid == IID_IApp)
	{
		IUnknown* p;
		if ((p = new CApp()) == NULL)
		{
			rc = E_OUTOFMEMORY;
		}
		else {
			rc = p->QueryInterface(iid, ppv);
			p->Release();
		}
	}
	return rc;
}

HRESULT __stdcall CMasterFactory::LockServer(BOOL bLock)
{
	if (bLock)
	{
		InterlockedIncrement((LONG*)&(g_ServerLocks));
	}
	else
	{
		InterlockedDecrement((LONG*)&(g_ServerLocks));
	}

	return S_OK;
}
