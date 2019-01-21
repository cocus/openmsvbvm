#pragma once
#include <ObjBase.h>
#include <objbase.h>
class CMasterFactory : public IClassFactory
{

public:
	CMasterFactory();
	~CMasterFactory();

	//interface IUnknown methods 
	HRESULT __stdcall QueryInterface(
		REFIID riid,
		void **ppObj);
	ULONG   __stdcall AddRef();
	ULONG   __stdcall Release();


	//interface IClassFactory methods 
	HRESULT __stdcall CreateInstance(IUnknown* pUnknownOuter,
		const IID& iid,
		void** ppv);
	HRESULT __stdcall LockServer(BOOL bLock);

private:
	long m_nRefCount;
};
/*
extern ULONG g_ServerLocks;     // блокировки сервера; 
template <typename T>
class ClassFactory :
	public IClassFactory //standart interface IClassFactory, provides creating components for outter requests
{
protected:
   // Reference count
   long          m_lRef;

public:
	ClassFactory(void);
	~ClassFactory(void);

	//IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,void  **ppv);
	virtual ULONG   STDMETHODCALLTYPE AddRef(void);
	virtual ULONG   STDMETHODCALLTYPE Release(void);

	
	virtual HRESULT STDMETHODCALLTYPE CreateInstance(LPUNKNOWN pUnk,  const IID& id, void** ppv); //creates component instance
	virtual HRESULT STDMETHODCALLTYPE LockServer (BOOL fLock); //обеспечивающий блокировку программы сервера в памяти. 
	//Блокируя сервер для других программ, клиент получает уверенность в том,
		//что доступ к нему будет быстро получен. Обычно это делается в целях повышения производительности.

};

template <typename T>
ClassFactory<T>::ClassFactory(void) : m_lRef (1) {
}

template <typename T>
ClassFactory<T>::~ClassFactory(void) {
}

//IUnknown
template <typename T>
HRESULT STDMETHODCALLTYPE ClassFactory<T>::QueryInterface(REFIID riid,void  **ppv) {
	HRESULT rc = S_OK;

	if (riid == IID_IUnknown) *ppv = (IUnknown*)this;
	else if (riid == IID_IClassFactory) *ppv = (IClassFactory*)this;
	else rc = E_NOINTERFACE;

	if (rc == S_OK)
		this->AddRef();

	return rc;
}

template <typename T>
ULONG   STDMETHODCALLTYPE ClassFactory<T>::AddRef(void) {
	InterlockedIncrement(&m_lRef);

	return this->m_lRef;
}

template <typename T>
ULONG   STDMETHODCALLTYPE ClassFactory<T>::Release(void) {
	InterlockedDecrement(&m_lRef);

	if ( this->m_lRef == 0) {
		delete this;
		return 0;
	}
	else
		return this->m_lRef;
}

template <typename T>
HRESULT STDMETHODCALLTYPE ClassFactory<T>::CreateInstance	(LPUNKNOWN pUnk, //basic interface IUnknown
																const IID& id,	// id of required interface
																void** ppv)	//required interface
{
	//SEQ;
	HRESULT rc = E_UNEXPECTED;

	if (pUnk != NULL) rc = CLASS_E_NOAGGREGATION;
	else {
		T* p;
		if ((p = new T()) == NULL) {
			rc = E_OUTOFMEMORY;
		}
		else {
			rc = p->QueryInterface(id,ppv);
			p->Release();
		}
	}
	return rc;
}

template <typename T>
HRESULT STDMETHODCALLTYPE ClassFactory<T>::LockServer (BOOL fLock) {
	if (fLock)  InterlockedIncrement ((LONG*)&(g_ServerLocks));
	else    InterlockedDecrement ((LONG*)&(g_ServerLocks));
	
	return S_OK;
}
*/