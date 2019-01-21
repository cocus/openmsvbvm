#include "vba_internal.h"
#include "vba_exception.h"

#include "vba_Locale.h"
#include "vba_objManipulation.h"
#include "DllObjectInterface.h"

#include "vba_structures.h"



typedef struct
{
	void					* lpVBVtable;
	unsigned int			lNull;
	class vbClassWrapper	* pWrapper;
} vba_VBVTable;




class vbClassWrapper : IDispatch
{
public:
	vbClassWrapper(vba_VBVTable * pWrapperVtable, TObjectInfoWithOptionals * pObjInfo);
	~vbClassWrapper();

	// IUnknown interface 
	HRESULT __stdcall QueryInterface(
		REFIID riid,
		void **ppObj
	);

	ULONG   __stdcall AddRef();
	ULONG   __stdcall Release();

	// IDispatch interface
	HRESULT __stdcall GetTypeInfoCount(
		UINT * pctInfo
	);

	HRESULT __stdcall GetTypeInfo(
		UINT itinfo,
		LCID lcid,
		ITypeInfo** pptinfo
	);

	HRESULT __stdcall GetIDsOfNames(
		REFIID riid,
		LPOLESTR* rgszNames,
		UINT cNames,
		LCID lcid,
		DISPID* rgdispid
	);
	
	HRESULT __stdcall Invoke(
		DISPID dispidMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS* pdispparams,
		VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo,
		UINT* puArgErr
	);

	// VB6 Init and Terminate invokers
	void InvokeVB6Initialize();
	void InvokeVB6Terminate();

private:
	long						m_nRefCount;   // for managing the reference count
	int							* m_pVtable = nullptr;
	vba_VBVTable				* m_pVBVTable = nullptr;
	TObjectInfoWithOptionals	* m_pObjInfo = nullptr;
};


#define DECLARE_BASIC_CLASS_WRAPPER(name, arg_types, arg_names, ret_type)			\
EXPORT ret_type __stdcall BASIC_CLASS_##name(										\
	arg_types																		\
)																					\
{																					\
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();											\
																					\
	DEBUG_WIDE(																		\
		"pFakeVtable %.8x",															\
		(unsigned int)pFakeVtable													\
	);																				\
																					\
	if (pFakeVtable->pWrapper)														\
	{																				\
		return pFakeVtable->pWrapper-> ## name(arg_names);							\
	}																				\
																					\
	return E_NOTIMPL;																\
}

#define DECLARE_BASIC_CLASS_TYPES_FOR_STRUCT(name, arg_types, arg_names, ret_type)	\
	ret_type(__stdcall * p ## name) (arg_types);

#define ASSIGN_MEMBERS_OF_BRIDGE_STRUCT(name, arg_types, arg_names, ret_type)		\
	bridgeStruct.p##name = BASIC_CLASS_##name;

#define BASIC_CLASS_WRAPPER_FUNCTIONS(decl)																											\
	decl(QueryInterface, ARGS(vba_VBVTable * pFakeVtable, REFIID riid, void ** ppObj), ARGS(riid, ppObj), ULONG)									\
	decl(AddRef, ARGS(vba_VBVTable * pFakeVtable), ARGS(), ULONG)																					\
	decl(Release, ARGS(vba_VBVTable * pFakeVtable), ARGS(), ULONG)																					\
	decl(GetTypeInfoCount, ARGS(vba_VBVTable * pFakeVtable, UINT * pctInfo), ARGS(pctInfo), HRESULT)												\
	decl(GetTypeInfo, ARGS(vba_VBVTable * pFakeVtable, UINT itinfo, LCID lcid, ITypeInfo** pptinfo), ARGS(itinfo, lcid, pptinfo), HRESULT)			\
	decl(GetIDsOfNames, ARGS(vba_VBVTable * pFakeVtable, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid), ARGS(riid, rgszNames, cNames, lcid, rgdispid), HRESULT) \
	decl(Invoke, ARGS(vba_VBVTable * pFakeVtable, DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr), ARGS(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr), HRESULT) \


/* This spawns the exported bridge functions to the BASIC_CLASS wrapper */
BASIC_CLASS_WRAPPER_FUNCTIONS(DECLARE_BASIC_CLASS_WRAPPER);

/* This structure holds pointer to all the BASIC_CLASS wrapper functions */
typedef struct
{
	BASIC_CLASS_WRAPPER_FUNCTIONS(DECLARE_BASIC_CLASS_TYPES_FOR_STRUCT);
} vba_BASIC_CLASS_IUnknownBridge;



ULONG __stdcall vbClassWrapper::AddRef()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ(
		"m_nRefCount was %.8x",
		m_nRefCount
	);

	return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall vbClassWrapper::Release()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ(
		"m_nRefCount was %.8x",
		m_nRefCount
	);

	static bool bDiscardCall = false;

	long nRefCount = 0;
	nRefCount = InterlockedDecrement(&m_nRefCount);
	if (nRefCount == 0 && bDiscardCall == false)
	{
		bDiscardCall = true;

		InvokeVB6Terminate(); // This will re-call Release, but we already set bDiscardCall

		delete this;
	}
	
	return nRefCount;
}

void vbClassWrapper::InvokeVB6Initialize()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	typedef void(__stdcall * pInit)(vba_VBVTable * ths);

	unsigned int ** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->aEventLinkArray) + ((this->m_pObjInfo->iEventCount - 2));

	if (pAddr == nullptr)
	{
		return;
	}

	pInit pCall = (pInit)(*pAddr);


	DEBUG_WIDE(
		"pCall %.8x",
		(unsigned int)pCall
	);

	pCall(this->m_pVBVTable);
}

void vbClassWrapper::InvokeVB6Terminate()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	typedef void(__stdcall * pInit)(vba_VBVTable * ths);

	unsigned int ** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->aEventLinkArray) + ((this->m_pObjInfo->iEventCount - 1));

	if (pAddr == nullptr)
	{
		return;
	}

	pInit pCall = (pInit)(*pAddr);

	DEBUG_WIDE(
		"pCall %.8x",
		(unsigned int)pCall
	);

	pCall(this->m_pVBVTable);
}

vbClassWrapper::vbClassWrapper(
	vba_VBVTable				*pWrapperVtable,
	TObjectInfoWithOptionals	*pObjInfo
)
	: m_nRefCount(1), m_pVBVTable(pWrapperVtable), m_pObjInfo(pObjInfo)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ(
		"New wrapper object! pWrapperVtable %.8x, pObjInfo %.8x",
		(unsigned int)pWrapperVtable,
		(unsigned int)pObjInfo
	);
}

vbClassWrapper::~vbClassWrapper()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	swprintf(
		debugW,
		DEBUG_STORAGE_SIZE - 1,
		L"vbClassWrapper::~vbClassWrapper(%.8x)!",
		(unsigned int)this
	);
	OutputDebugStringW(debugW);

	if (m_pVBVTable)
	{
		if (m_pVBVTable->lpVBVtable)
		{
			free(m_pVBVTable->lpVBVtable);
			m_pVBVTable->lpVBVtable = nullptr;
		}

		free(m_pVBVTable);
		m_pVBVTable = nullptr;
	}
}

HRESULT __stdcall vbClassWrapper::QueryInterface(
	REFIID riid,
	void ** ppObj
)
{
	return E_NOTIMPL;
}

HRESULT __stdcall vbClassWrapper::GetTypeInfoCount(
	UINT * pctInfo
)
{
	return E_NOTIMPL;
}

HRESULT __stdcall vbClassWrapper::GetTypeInfo(
	UINT itinfo,
	LCID lcid,
	ITypeInfo ** pptinfo
)
{
	return E_NOTIMPL;
}

HRESULT __stdcall vbClassWrapper::GetIDsOfNames(
	REFIID riid,
	LPOLESTR * rgszNames,
	UINT cNames,
	LCID lcid,
	DISPID * rgdispid
)
{
	return E_NOTIMPL;
}

HRESULT __stdcall vbClassWrapper::Invoke(
	DISPID dispidMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS * pdispparams,
	VARIANT * pvarResult,
	EXCEPINFO * pexcepinfo,
	UINT * puArgErr
)
{
	return E_NOTIMPL;
}


/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT void __stdcall __vbaHresultCheckObj(
	int				arg1,
	int				arg2,
	GUID			*arg3,
	__int16			arg4
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	DEBUG_ASCII(
		"arg1 %.8x, arg2 %.8x, arg3 %.8x, arg4 %.8x",
		(unsigned int)arg1,
		(unsigned int)arg2,
		(unsigned int)arg3,
		(unsigned int)arg4
	);
}

extern GUID CLSID_VBGlobal;
extern GUID CLSID_App;

/**
 * @brief			Instantiates a new COM object via it's CoClass and Interface, if specified.
 * @param			pvbNewData		Pointer to a VB defined structure with info about the object.
 * @param			ppv				Pointer to an IUnknown pointer, where the created object will be stored.
 * @returns			HRESULT of the operation.
 */
EXPORT HRESULT __stdcall __vbaNew2(
	vba_new_data_arg_t		* pvbNewData,
	IUnknown				** ppv
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvbNewData %.8x, reval %.8x",
		(unsigned int)pvbNewData,
		(unsigned int)ppv
	);

	if (pvbNewData->lpguidCoClass)
	{
		DEBUG_WIDE_GUID(*pvbNewData->lpguidCoClass, "CoClass");
	}

	if (pvbNewData->lpguidInterface)
	{
		DEBUG_WIDE_GUID(*pvbNewData->lpguidInterface, "Interface");
	}

	HRESULT hr;

	/* Check if the required CLSID is one of the internal VB CLSIDs */
	if (*pvbNewData->lpguidCoClass == CLSID_VBGlobal)
	{
		DEBUG_WIDE(
			"Got VBGlobal CoClass, using internal class"
		);

		hr = DllGetClassObject(
			*pvbNewData->lpguidCoClass,
			*pvbNewData->lpguidInterface,
			(LPVOID*)ppv
		);

		DEBUG_WIDE(
			"DllGetClassObject ret %.8x",
			hr
		);

		if (hr != S_OK)
		{
			vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
			return hr;
		}
	}
	else
	{
		/* Create instance from the specified CoClass GUID */
		hr = CoCreateInstance(
			*pvbNewData->lpguidCoClass,
			nullptr,
			CLSCTX_INPROC_SERVER,
			*pvbNewData->lpguidInterface,
			(LPVOID*)ppv
		);

		DEBUG_WIDE(
			"CoCreateInstance ret %.8x",
			hr
		);

		if (hr != S_OK)
		{
			vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
			return hr;
		}
	}

	return hr;
}

/**
 * @brief			Constructs a wrapper object for a VB6 class, and instantiates that class.
 * @param			pvbNewData			Object info pointer.
 * @returns			Valid vba_VBVTable pointer on success, nullptr otherwise.
 */
EXPORT vba_VBVTable * __stdcall __vbaNew(
	TObjectInfoWithOptionals	* pvbNewData
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (pvbNewData == nullptr)
	{
		DEBUG_WIDE(
			"pvbNewData %.8x",
			(unsigned int)pvbNewData
		);

		return nullptr;
	}

	DEBUG_WIDE(
		"pvbNewData %.8x, pvbNewData->iEventCount %.8x, pvbNewData->hdr.aObjectHeader %.8x, pvbNewData->aBasicClassObject %.8x",
		(unsigned int)pvbNewData,
		(unsigned int)pvbNewData->iEventCount,
		(unsigned int)pvbNewData->hdr.aObjectHeader,
		(unsigned int)pvbNewData->aBasicClassObject
	);

	if (!pvbNewData->hdr.aObjectHeader)
	{
		return nullptr;
	}

	TObject * tObj = (TObject*)(pvbNewData->hdr.aObjectHeader);
	DEBUG_WIDE(
		"tObj->aModulePublic %.8x, tObj->aModuleStatic %.8x, tObj->aPublicBytes %.8x, tObj->aStaticBytes %.8x, tObj->lMethodCount %.8x, tObj->oStaticVars %.8x",
		(unsigned int)tObj->aModulePublic,
		(unsigned int)tObj->aModuleStatic,
		(unsigned int)tObj->aPublicBytes,
		(unsigned int)tObj->aStaticBytes,
		(unsigned int)tObj->lMethodCount,
		(unsigned int)tObj->oStaticVars
	);

	if (tObj->aPublicBytes == nullptr)
	{
		return nullptr;
	}

	DEBUG_WIDE(
		"tObj->aPublicBytes->iConst1 %.8x, tObj->aPublicBytes->iSize %.8x",
		tObj->aPublicBytes->iConst1,
		tObj->aPublicBytes->iSize
	);

	UINT uiVtableCount = pvbNewData->iEventCount;

	/* pWrapperVTable is technically the "VB class" with it's functions after the IDispatch stuff */
	void * pWrapperVtable = malloc(tObj->aPublicBytes->iSize);//sizeof(void*) * uiVtableCount);
	if (pWrapperVtable == nullptr)
	{
		return nullptr;
	}
	memset(
		pWrapperVtable,
		0,
		tObj->aPublicBytes->iSize//sizeof(void*) * uiVtableCount
	);

	/* Setup the bridgeStruct (IUnk-like that bridges VB's VTable to COM) */
	vba_BASIC_CLASS_IUnknownBridge bridgeStruct;
	BASIC_CLASS_WRAPPER_FUNCTIONS(ASSIGN_MEMBERS_OF_BRIDGE_STRUCT);
	
	/* And copy it to the pWrapperVTable */
	memcpy(
		pWrapperVtable,
		&bridgeStruct,
		sizeof(vba_BASIC_CLASS_IUnknownBridge)
	);

	/* Copy the VB Specified vtable to the pWrapperVTable after bridgeStruct */
	memcpy(
		(void*)((unsigned int)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge)),
		pvbNewData->aEventLinkArray,
		sizeof(void*) * uiVtableCount
	);

	/* Setup a vba_VBVTable struct, which is what we'll return and it'll be what VB code will use as 'this' (i.e. also contains the local storage of the class) */
	vba_VBVTable * ret = (vba_VBVTable*)malloc(sizeof(vba_VBVTable) + tObj->aPublicBytes->iSize);
	if (ret == nullptr)
	{
		free(pWrapperVtable);
		return nullptr;
	}
	memset(
		ret,
		0,
		sizeof(vba_VBVTable) + tObj->aPublicBytes->iSize
	);

	ret->lpVBVtable = pWrapperVtable;


	/* Create a vbClassWrapper object, and assign the vba_VBVTable pointer to it */
	ret->pWrapper = new vbClassWrapper(ret, pvbNewData);

	DEBUG_WIDE(
		"ret %.8x, ret->lpVBVtable %.8x spans to %.8x, ret->pWrapper %.8x",
		(unsigned int)ret,
		(unsigned int)ret->lpVBVtable,
		(unsigned int)ret->lpVBVtable + tObj->aPublicBytes->iSize,
		(unsigned int)ret->pWrapper		
	);

	if (ret->pWrapper == nullptr)
	{
		free(pWrapperVtable);
		free(ret);
		return nullptr;
	}

	/* Call the Initialize method */
	DEBUG_WIDE(
		"pvbNewData->oInitializeEvent %.8x, pvbNewData->oTerminateEvent %.8x, pvbNewData->aEventLinkArray %.8x, pvbNewData->iEventCount %.8x",
		pvbNewData->oInitializeEvent,
		pvbNewData->oTerminateEvent,
		(unsigned int)(void*)pvbNewData->aEventLinkArray,
		pvbNewData->iEventCount
	);

	ret->pWrapper->InvokeVB6Initialize();

	return ret;
}








/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT  __stdcall EVENT_SINK_AddRef(
	void			*arg1
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"arg1 %.8x",
		(unsigned int)arg1
	);

	return E_NOTIMPL;
}

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT  __stdcall EVENT_SINK_Release(
	void			*arg1
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"arg1 %.8x",
		(unsigned int)arg1
	);

	return E_NOTIMPL;
}

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT  __stdcall EVENT_SINK_QueryInterface(
	void			*arg1,
	REFIID			arg2,
	void			**ppvObject
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"arg1 %.8x, arg2 %.8x, ppvObject %.8x",
		(unsigned int)arg1,
		(unsigned int)&arg2,
		(unsigned int)ppvObject
	);

	return E_NOTIMPL;
}

/**
 * @brief			Releases the ref of the destination IUnknown pointer (if not null), and sets it
 *					with the source pointer.
 * @param			ppiuDest		Pointer to a IUnknown *. This is where the piuSrc value will be written.
 * @param			piuSrc			Pointer to a IUnknown.
 * @returns			*ppiuDest always.
 */
EXPORT IUnknown * __stdcall __vbaObjSet(
	IUnknown		**ppiuDest,
	IUnknown		*piuSrc
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ppiuDest %.8x, piuSrc %.8x",
		(unsigned int)ppiuDest,
		(unsigned int)piuSrc
	);

	if (*ppiuDest)
	{
		(*ppiuDest)->Release();
	}

	*ppiuDest = piuSrc;

	return *ppiuDest;
}

/**
 * @brief			Releases the ref of the destination IUnknown pointer (if not null), and sets it
 *					with the source pointer. Also adds a reference of the source IUnknown object.
 * @param			ppiuDest		Pointer to a IUnknown *. This is where the piuSrc value will be written.
 * @param			piuSrc			Pointer to a IUnknown.
 * @returns			*ppiuDest always.
 */
EXPORT IUnknown * __stdcall __vbaObjSetAddref(
	IUnknown		**ppiuDest,
	IUnknown		*piuSrc
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ppiuDest %.8x, piuSrc %.8x",
		(unsigned int)ppiuDest,
		(unsigned int)piuSrc
	);

	if (piuSrc)
	{
		piuSrc->AddRef();
	}

	return __vbaObjSet(ppiuDest, piuSrc);
}

/**
 * @brief			Gets a pointer to a specified interface from an IUnknown (if exists).
 * @param			piunkIn			Pointer to a IUnknown.
 * @param			riid			IID of the interface to obtain.
 * @returns			Valid pointer to an interface on success, nullptr otherwise.
 */
EXPORT void * __stdcall __vbaCastObj(
	IUnknown		*piunkIn,
	REFIID			riid
)
{
	HRESULT			hr;
	void			*ppvObject = nullptr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"piunkIn %.8x, riid %.8x",
		(unsigned int)piunkIn,
		(unsigned int)(*(unsigned int*)&riid)
	);

	if (!piunkIn)
	{
		/* We don't have a IUnknown pointer, so ... */
		vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR); // TODO: Check this exception
	}
	else
	{
		hr = piunkIn->QueryInterface(riid, &ppvObject);
		
		/* Check for return values */
		if (hr != S_OK)
		{
			/* QueryInterface returned an error */
			switch (hr)
			{
				case E_NOINTERFACE:
				{
					vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH); // TODO: Check this exception
					break;
				}

				default:
				{
					vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check this exception
				}
			}
		}
		else if (!ppvObject)
		{
			/* QueryInterface didn't set the ppvObject */
			vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET); // TODO: Check this exception
		}
	}

	return ppvObject;
}

/**
 * @brief			Creates a COM object from a specified class name.
 * @param			pvargObject		Created object will be set into this Variant variable.
 * @param			bstrClassName	Class name of the object.
 * @param			bstrServerName	Server name where the object will be created.
 * @returns			HRESULT of the operation.
 */
EXPORT HRESULT __stdcall rtcCreateObject2(
	VARIANTARG		*pvargObject,
	BSTR			bstrClassName,
	BSTR			bstrServerName
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvargObject %.8x, bstrClassName %.8x ('%ls'), bstrServerName %.8x",
		(unsigned int)pvargObject,
		(unsigned int)bstrClassName,
		bstrClassName,
		(unsigned int)bstrServerName
	);

	HRESULT			hr;
	CLSID			rCLSID;
	
	hr = CLSIDFromProgIDEx(bstrClassName, &rCLSID);

	DEBUG_WIDE_GUID(
		rCLSID,
		"CLSIDFromProgIDEx"
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	IUnknown		*ppv;
	hr = CoCreateInstance(
		rCLSID,
		nullptr,
		CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
		IID_IUnknown,
		(LPVOID*)&ppv
	);
	DEBUG_WIDE(
		"CoCreateInstance = %.8x, ppv = %.8x",
		(unsigned int)hr,
		(unsigned int)ppv
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	IDispatch		*disp;
	hr = ppv->QueryInterface(
		IID_IDispatch,
		(LPVOID*)&disp
	);
	DEBUG_WIDE(
		"ppv->QueryInterface = %.8x, disp = %.8x",
		(unsigned int)hr,
		(unsigned int)disp
	);

	if (hr != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	pvargObject->vt = VT_DISPATCH;
	pvargObject->pdispVal = disp;

	return hr;
}

/**
 * @brief			Returns a pointer to an IDispatch from a VARIANTARG.
 * @param			pvargIn			Pointer to the VARIANTARG where the IDispatch will be extracted from.
 * @returns			plVal from the VARIANTARG argument, only if the object type is VT_DISPATCH. 0 otherwise.
 */
EXPORT IDispatch * __stdcall __vbaObjVar(
	VARIANTARG		* pvargIn
)
{
	VARTYPE				vtype;
	IDispatch			* ret = nullptr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvargIn %.8x",
		(unsigned int)pvargIn
	);

	if (!pvargIn)
	{
		vbaRaiseException(VBA_EXCEPTION_OBJECT_REQUIRED);
		return nullptr;
	}

	DEBUG_WIDE(
		"pvargIn type %.8x, pvargIn plVal %.8x",
		(unsigned int)pvargIn->vt,
		(unsigned int)pvargIn->plVal
	);

	vtype = pvargIn->vt;
	ret = (IDispatch*)pvargIn->plVal;

	/* Get the true pointer if it's a BYREF */
	if (vtype & VT_BYREF)
	{
		/* De-ref the pointer */
		ret = *(IDispatch**)ret;

		/* Keep the VT_ type only and ditch the other flags */
		vtype &= VT_TYPEMASK;
	}
	
	/* If the object is not a dispatch, raise an exception */
	if (vtype != VT_DISPATCH)
	{
		vbaRaiseException(VBA_EXCEPTION_OBJECT_REQUIRED);
		return nullptr;
	}
	return ret;
}

/**
 * @brief			Invokes a method of an IDispatch object, with no return value.
 * @param			pidObject		Pointer to a Variant where the return value will be stored.
 * @param			bstrMethodName	Name of the method to invoke.
 * @param			argCount		Number of arguments used in the variadic argument.
 * @param			...				Arguments.
 */
EXPORT void __cdecl __vbaLateMemCall(
	IDispatch		* pidObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
)
{
	HRESULT			hr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pidObject %.8x, bstrMethodName %.8x, argCount %.8x",
		(unsigned int)pidObject,
		(unsigned int)bstrMethodName,
		argCount
	);

	if (!pidObject)
	{
		vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
		return;
	}

	DISPID			dispID;

	/* Get the ID of the method */
	hr = pidObject->GetIDsOfNames(
		IID_NULL,
		&bstrMethodName,
		1,
		LOCALE_USER_DEFAULT,
		&dispID
	);

	DEBUG_WIDE(
		"GetIDsOfNames = %.8x, dispID = %.8x",
		(unsigned int)hr,
		dispID
	);

	if (hr != S_OK)
	{
		DEBUG_WIDE(
			"GetIDsOfNames failed! GetLastError() = %.8x",
			GetLastError()
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right

		return;
	}

	va_list args;
	va_start(args, argCount);

	DISPPARAMS		dispParamsInput;

	dispParamsInput.cArgs = argCount;
	dispParamsInput.rgvarg = (VARIANTARG*)args;

	va_end(args);

	dispParamsInput.cNamedArgs = 0;
	dispParamsInput.rgdispidNamedArgs = nullptr;

	EXCEPINFO		excepInfo;
	UINT			uArgErr;

	/* Invoke the method */
	hr = pidObject->Invoke(
		dispID,
		IID_NULL,
		LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYGET | DISPATCH_METHOD,
		&dispParamsInput,
		nullptr,
		&excepInfo,
		&uArgErr
	);

	DEBUG_WIDE(
		"pidObject->Invoke = %.8x",
		(unsigned int)hr
	);

	if (hr != S_OK)
	{
		DEBUG_WIDE(
			"pidObject->Invoke failed! GetLastError() = %.8x, excepInfo = '%ls', uArgErr = %.8x",
			GetLastError(),
			excepInfo.bstrDescription,
			uArgErr
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right
	}
}

/**
 * @brief			Invokes a method of a Variant object, and returns the return value of the invoke.
 * @param			pvargRet		Pointer to a Variant where the return value will be stored.
 * @param			pvargObject		Variant where the object is stored.
 * @param			bstrMethodName	Name of the method to invoke.
 * @param			argCount		Number of arguments used in the variadic argument.
 * @param			...				Arguments.
 * @returns			pvargRet always.
 */
EXPORT VARIANTARG * __cdecl __vbaVarLateMemCallLdRf(
	VARIANTARG		* pvargRet,
	VARIANTARG		* pvargObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
)
{
	HRESULT			hr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvargRet %.8x, pvargRet->vt %.8x, pvargObject %.8x, bstrMethodName %.8x, argCount %.8x",
		(unsigned int)pvargRet,
		(unsigned int)pvargRet->vt,
		(unsigned int)pvargObject,
		(unsigned int)bstrMethodName,
		argCount
	);

	if (!pvargObject)
	{
		vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
		return pvargRet;
	}

	IDispatch		*pidObject;

	/* Get the IDispatch object from the Variant */
	pidObject = __vbaObjVar(pvargObject);

	if (!pidObject)
	{
		DEBUG_WIDE(
			"Could not get IDispatch object from the pvargObject %.8x",
			(unsigned int)pvargObject
		);

		return pvargRet;
	}

	DISPID			dispID;

	/* Get the ID number of the requested method */
	hr = pidObject->GetIDsOfNames(
		IID_NULL,
		&bstrMethodName,
		1,
		LOCALE_USER_DEFAULT,
		&dispID
	);

	DEBUG_WIDE(
		"GetIDsOfNames = %.8x, dispID = %.8x",
		(unsigned int)hr,
		dispID
	);

	if (hr != S_OK)
	{
		DEBUG_WIDE(
			"GetIDsOfNames failed! GetLastError() = %.8x",
			GetLastError()
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right

		return pvargRet;
	}

	va_list args;
	va_start(args, argCount);

	DISPPARAMS		dispParamsInput;

	dispParamsInput.cArgs = argCount;
	dispParamsInput.rgvarg = (VARIANTARG*)args;

	DEBUG_WIDE(
		"dispParamsInput.rgvarg %.8x",
		(unsigned int)dispParamsInput.rgvarg
	);

	va_end(args);

	dispParamsInput.cNamedArgs = 0;
	dispParamsInput.rgdispidNamedArgs = nullptr;

	EXCEPINFO		excepInfo;
	UINT			uArgErr;

	/* Invoke the method */
	hr = pidObject->Invoke(
		dispID,
		IID_NULL,
		LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYGET | DISPATCH_METHOD,
		&dispParamsInput,
		pvargRet,
		&excepInfo,
		&uArgErr
	);

	DEBUG_WIDE(
		"pidObject->Invoke = %.8x, pvargRet->vt %.8x",
		(unsigned int)hr,
		(unsigned int)pvargRet->vt
	);

	if (hr != S_OK)
	{
		DEBUG_WIDE(
			"pidObject->Invoke failed! GetLastError() = %.8x, excepInfo = '%ls', uArgErr = %.8x",
			GetLastError(),
			excepInfo.bstrDescription,
			uArgErr
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right
	}

	return pvargRet;
}

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
HRESULT objIDispatchGetDefaultValue(
	IDispatch		*pidObject,
	VARIANTARG		*pvargValueOut
)
{
	HRESULT			hr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pidObject %.8x, pvargValueOut %.8x",
		(unsigned int)pidObject,
		(unsigned int)pvargValueOut
	);

	if (!pidObject)
	{
		vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
		return -1;
	}

	DISPPARAMS		dispParamsInput = { 0 };
	EXCEPINFO		excepInfo;
	UINT			uArgErr;

	/* Invoke the method */
	hr = pidObject->Invoke(
		0,
		IID_NULL,
		LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYGET | DISPATCH_METHOD,
		&dispParamsInput,
		pvargValueOut,
		&excepInfo,
		&uArgErr
	);

	DEBUG_WIDE(
		"pidObject->Invoke = %.8x, pvargValueOut->vt %.8x",
		(unsigned int)hr,
		(unsigned int)pvargValueOut->vt
	);

	if (hr != S_OK)
	{
		DEBUG_WIDE(
			"pidObject->Invoke failed! GetLastError() = %.8x, excepInfo = '%ls', uArgErr = %.8x",
			GetLastError(),
			excepInfo.bstrDescription,
			uArgErr
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right
	}

	return hr;
}

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
EXPORT VARIANTARG * __cdecl __vbaVarLateMemCallLd(
	VARIANTARG		* pvargRet,
	VARIANTARG		* pvarObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"bridgeing to __vbaVarLateMemCallLdRf but it's wrong and will crash!"
	);

	va_list args;
	va_start(args, argCount);
	return __vbaVarLateMemCallLdRf(pvargRet, pvarObject, bstrMethodName, argCount, args);
}

/**
 * @brief			Releases a COM Object (IUnknown) via its pointer, and nulls it.
 * @param			ppunkObj		Pointer to an object to release, and null the pointer.
 * @returns			Dereferenced object of ppunkObj.
 */
EXPORT IUnknown * __fastcall __vbaFreeObj(
	IUnknown		** punkObj
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	DEBUG_ASCII(
		"punkObj %.8x",
		(unsigned int)punkObj
	);

	if (punkObj)
	{
		DEBUG_ASCII(
			"*punkObj %.8x",
			(unsigned int)*((unsigned int*)punkObj)
		);

		if (*((unsigned int*)punkObj))
		{
			(*punkObj)->Release();
			*(unsigned int*)punkObj = 0;
		}
	}

	return *punkObj;
}

/**
 * @brief			Releases a list of COM Objects (IUnknowns) via their pointers, and nulls them.
 * @param			argCount		Count of elements.
 * @param			...				Pointers to objects to release and null them.
 */
EXPORT void __cdecl __vbaFreeObjList(
	unsigned int	uiArgCount,
	...
)
{
	IUnknown	**piunkElement;

	va_list		args;
	va_start(args, uiArgCount);

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	while (uiArgCount--)
	{
		piunkElement = va_arg(args, IUnknown**);
		DEBUG_WIDE(
			"arg() %.8x, argsRemaining %.8x",
			(unsigned int)piunkElement,
			uiArgCount
		);

		__vbaFreeObj(piunkElement);
	}

	va_end(args);
}