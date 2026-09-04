#include "vba_internal.h"
#include "Exceptions.hpp"
#include <wctype.h>

#include "vba_Locale.h"
#include "ObjectManipulation.hpp"
#include "DllObjectInterface.hpp"

#include "vba_structures.h"
#include "EventDispatch.hpp"
#include "FormWindow.hpp"
#include "FormProperties.hpp"
#include "ClassWrapper.hpp"
#include "FormWrapper.hpp"


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

	/* bWInitializeEvent/bWTerminateEvent are byte offsets measured from the wEventCount field
	   itself (8 bytes before lpEvents), NOT plain indices into lpEvents and NOT always the last
	   two entries -- a class with WithEvents members appends its Get/Put/Set accessor thunks
	   after Class_Initialize/Class_Terminate, so "last two slots" only works by coincidence for
	   classes without WithEvents. Confirmed via msvbvm60 disassembly + a real WithEvents test case. */
	unsigned int index = ((unsigned int)this->m_pObjInfo->opt.bWInitializeEvent - 8) / sizeof(void*);
	unsigned int ** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->opt.lpEvents) + index;

	if (pAddr == nullptr)
	{
		return;
	}

	pInit pCall = (pInit)(*pAddr);


	DEBUG_WIDE(
		"pCall %.8x",
		(unsigned int)pCall
	);

	// Not every class implements Class_Initialize -- lpEvents can hold a null slot
	// (or the index can land outside the populated part of the table for classes
	// whose event/WithEvents layout differs, e.g. Forms). Calling through that is
	// an unconditional crash, so skip when there's nothing to call.
	if (pCall == nullptr)
	{
		return;
	}

	CurrentInstanceScope instanceScope(this->m_pVBVTable);
	pCall(this->m_pVBVTable);
}

void vbClassWrapper::InvokeVB6Terminate()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	typedef void(__stdcall * pInit)(vba_VBVTable * ths);

	unsigned int index = ((unsigned int)this->m_pObjInfo->opt.bWTerminateEvent - 8) / sizeof(void*);
	unsigned int ** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->opt.lpEvents) + index;

	if (pAddr == nullptr)
	{
		return;
	}

	pInit pCall = (pInit)(*pAddr);

	DEBUG_WIDE(
		"pCall %.8x",
		(unsigned int)pCall
	);

	// See the matching guard in InvokeVB6Initialize -- not every class implements
	// Class_Terminate, so a null slot here is expected and must not be called.
	if (pCall == nullptr)
	{
		return;
	}

	CurrentInstanceScope instanceScope(this->m_pVBVTable);
	pCall(this->m_pVBVTable);
}

vbClassWrapper::vbClassWrapper(
	vba_VBVTable				*pWrapperVtable,
	ObjectInfoWithOptional*pObjInfo
)
	: m_nRefCount(1), m_pVBVTable(pWrapperVtable), m_pObjInfo(pObjInfo), m_connectionPoint(static_cast<IDispatch*>(this))
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
	if (!ppObj)
	{
		return E_POINTER;
	}

	if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
	{
		/* m_pVBVTable is vtable-first (lpVBVtable is its first field), i.e. it's
		   itself the valid IDispatch-shaped pointer compiled VB6 code and other
		   parts of this project already treat as "the object". */
		*ppObj = m_pVBVTable;
		AddRef();
		return S_OK;
	}

	if (IsEqualIID(riid, IID_IConnectionPointContainer))
	{
		*ppObj = static_cast<IConnectionPointContainer*>(this);
		AddRef();
		return S_OK;
	}

	*ppObj = nullptr;
	return E_NOINTERFACE;
}

HRESULT __stdcall vbClassWrapper::GetTypeInfoCount(
	UINT * pctInfo
)
{
	if (pctInfo)
	{
		*pctInfo = 0; /* no real ITypeInfo object provided yet */
	}
	return S_OK;
}

HRESULT __stdcall vbClassWrapper::GetTypeInfo(
	UINT itinfo,
	LCID lcid,
	ITypeInfo ** pptinfo
)
{
	return TYPE_E_ELEMENTNOTFOUND;
}

HRESULT __stdcall vbClassWrapper::GetIDsOfNames(
	REFIID riid,
	LPOLESTR * rgszNames,
	UINT cNames,
	LCID lcid,
	DISPID * rgdispid
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (!m_pObjInfo || !m_pObjInfo->hdr.lpObject || !rgszNames || !rgdispid)
	{
		return E_NOTIMPL;
	}

	/* Real name -> DISPID resolution sourced from this compiled class's own
	   PublicObjectDescriptor.lpMethodNames table (already-populated for native
	   builds, including for Private WithEvents handler subs. */
	PublicObjectDescriptor * pDesc = m_pObjInfo->hdr.lpObject;
	HRESULT hrOverall = S_OK;

	for (UINT i = 0; i < cNames; i++)
	{
		rgdispid[i] = DISPID_UNKNOWN;

		for (DWORD m = 0; pDesc->lpMethodNames && m < pDesc->dwMethodCount; m++)
		{
			LPCSTR pszName = pDesc->lpMethodNames[m];
			if (!pszName)
			{
				continue;
			}

			LPCSTR pA = pszName;
			LPOLESTR pW = rgszNames[i];
			bool matches = true;
			while (*pA && *pW)
			{
				if (towlower((wchar_t)(unsigned char)*pA) != towlower(*pW))
				{
					matches = false;
					break;
				}
				pA++;
				pW++;
			}
			matches = matches && (*pA == 0) && (*pW == 0);

			if (matches)
			{
				rgdispid[i] = (DISPID)(m + 1);
				break;
			}
		}

		if (rgdispid[i] == DISPID_UNKNOWN)
		{
			hrOverall = DISP_E_UNKNOWNNAME;
		}
	}

	DEBUG_WIDE(
		"cNames %.8x, hr %.8x",
		cNames,
		(unsigned int)hrOverall
	);

	return hrOverall;
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

HRESULT __stdcall vbClassWrapper::EnumConnectionPoints(
	IEnumConnectionPoints ** ppEnum
)
{
	return E_NOTIMPL;
}

HRESULT __stdcall vbClassWrapper::FindConnectionPoint(
	REFIID riid,
	IConnectionPoint ** ppCP
)
{
	if (!ppCP)
	{
		return E_POINTER;
	}

	/* This class only ever exposes one outgoing (source) interface, so riid isn't
	   disambiguated -- see the plan's Known Limitations for a class with more than
	   one distinct outgoing interface. */
	*ppCP = &m_connectionPoint;
	m_connectionPoint.AddRef();
	return S_OK;
}

void vbClassWrapper::RaiseEvent(DISPID dispId, VARIANTARG * pArgs, DWORD argCount)
{
	RaiseEventOnSinks(m_connectionPoint.Sinks(), dispId, pArgs, argCount);
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

	// arg1 is the HRESULT returned by the object call this guards (e.g. Load/Unload
	// on the VB global object). This was a no-op stub, which let compiled code march
	// on as if a failed call (e.g. our still-E_NOTIMPL VBGlobalImpl::Load) had
	// succeeded, crashing later on unrelated state. Raise the matching VB runtime
	// error instead, same as every other HRESULT-checking call site in this project.
	HRESULT hr = (HRESULT)arg1;
	if (FAILED(hr))
	{
		vbaRaiseException(vbaErrorFromHRESULT(hr));
	}
} /* __vbaHresultCheckObj */

/**
 * @brief			Returns true if pDesc describes a Form- or MDIForm-derived class.
 *
 * A first attempt at this compared opt.lpuuidObjectTypes[0] against a GUID literal
 * captured from one specific Form1 build -- that turned out to be WRONG: re-checking
 * across two separate compiles of the same Form1 showed that GUID is freshly
 * randomized on every single compile (same as opt.clsidObjectClass), not a stable
 * "this is a Form" marker at all. That bug silently made this function return false
 * after any rebuild, skipping the vtable padding entirely and leaving `.Show`'s compiled
 * call through `[lpVBVtable + 0x2B0]` reading unrelated heap memory.
 *
 * PublicObjectDescriptor.fObjectType (see that field's own doc comment in
 * vba_structures.h for the full confirmed bit table) is the real, stable signal: bit
 * 0x80 is the discriminator, confirmed shared by both a plain Form and an MDIForm
 * (dumped from a whole project's ObjectTable.lpObjectArray, live, cross-checked
 * against VBHeader.wFormCount -- see that struct's own lpGuiTable comment for why
 * that field looked like a promising shortcut for this but wasn't one).
 * UserControl/UserDocument/PropertyPage aren't covered -- this project doesn't
 * support those object kinds at all yet, so whether they'd need to be treated as
 * "form-like" here too hasn't come up.
 */
static bool IsFormLikeDescriptor(
	ObjectInfoWithOptional	*pDesc
)
{
	if (!pDesc->hdr.lpObject)
	{
		return false;
	}

	return (pDesc->hdr.lpObject->fObjectType & 0x80) != 0;
}

/**
 * @brief			Constructs a wrapper object for a VB6 class, and instantiates that class.
 * @param			pvbNewData			Object info pointer.
 * @returns			Valid vba_VBVTable pointer on success, nullptr otherwise.
 */
EXPORT vba_VBVTable * __stdcall __vbaNew(
	ObjectInfoWithOptional* pvbNewData
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
		"pvbNewData %.8x, pvbNewData->iEventCount %.8x, pvbNewData->hdr.lpObject %.8x, pvbNewData->opt.lpBasicClassObject %.8x",
		(unsigned int)pvbNewData,
		(unsigned int)pvbNewData->opt.wEventCount,
		(unsigned int)pvbNewData->hdr.lpObject,
		(unsigned int)pvbNewData->opt.lpBasicClassObject
	);

	if (!pvbNewData->hdr.lpObject)
	{
		return nullptr;
	}

	PublicObjectDescriptor * tObj = static_cast<PublicObjectDescriptor*>(pvbNewData->hdr.lpObject);
	DEBUG_WIDE(
		"tObj->lpModulePublic %.8x, tObj->lpModuleStatic %.8x, tObj->lpPublicBytes %.8x, tObj->lpStaticBytes %.8x, tObj->dwMethodCount %.8x, tObj->bStaticVars %.8x",
		(unsigned int)tObj->lpModulePublic,
		(unsigned int)tObj->lpModuleStatic,
		(unsigned int)tObj->lpPublicBytes,
		(unsigned int)tObj->lpStaticBytes,
		(unsigned int)tObj->dwMethodCount,
		(unsigned int)tObj->bStaticVars
	);

	if (tObj->lpPublicBytes == nullptr)
	{
		return nullptr;
	}

	DEBUG_WIDE(
		"tObj->lpPublicBytes->iConst1 %.8x, tObj->lpPublicBytes->iSize %.8x",
		tObj->lpPublicBytes->iConst1,
		tObj->lpPublicBytes->iSize
	);

	UINT uiVtableCount = pvbNewData->opt.wEventCount;
	bool bIsFormLike = IsFormLikeDescriptor(pvbNewData);

	/* tObj->lpPublicBytes->iSize is the size of the class's OWN public/module-level variable
	   storage (used below for the per-instance data appended after vba_VBVTable) -- it has
	   nothing to do with the vtable blob's size, and can legitimately be 0 for a class with no
	   public variables (e.g. clsTestClass1). The vtable blob's size is the bridge plus the
	   class's own compiled method/event table, computed independently here. */
	size_t vtableBlobSize = sizeof(vba_BASIC_CLASS_IUnknownBridge) + (size_t)sizeof(void*) * uiVtableCount;

	/* Form-derived classes compile direct vtable calls (e.g. Show) at large, fixed byte
	   offsets into this SAME per-instance vtable -- confirmed live: `f.Show vbModal`
	   compiles to `call dword ptr [f->lpVBVtable + 0x2B0]`. Pad the blob out to fit
	   the real _Form interface's own tail (everything after IDispatch -- see
	   GetFormInterfaceVtableTail, FormWrapper.cpp/.hpp) so every one of its slots is
	   valid, inserted *before* this class's own wEventCount-driven slots (which the
	   compiler lays out immediately after however large its own vtable actually is).
	   Only Show is a real implementation in that copied tail; every other intrinsic
	   slot is an inert stub -- not yet reverse-engineered, and not called by any test
	   exercised so far (see the plan's Known Limitations). */
	size_t formVtableTailSlotCount = 0;
	void * const *pFormVtableTail = nullptr;
	size_t formIntrinsicPadding = 0;
	if (bIsFormLike)
	{
		pFormVtableTail = GetFormInterfaceVtableTail(&formVtableTailSlotCount);
		formIntrinsicPadding = formVtableTailSlotCount * sizeof(void*);
		vtableBlobSize += formIntrinsicPadding;
	}

	/* pWrapperVTable is technically the "VB class" with it's functions after the IDispatch stuff */
	void * pWrapperVtable = malloc(vtableBlobSize);
	if (pWrapperVtable == nullptr)
	{
		return nullptr;
	}
	memset(
		pWrapperVtable,
		0,
		vtableBlobSize
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

	if (bIsFormLike && pFormVtableTail)
	{
		/* Landing every real _Form member (including Show, at its true compiler-
		   assigned slot -- see GetFormInterfaceVtableTail's own comment for why that
		   lines up with the real, disassembly-confirmed 0x2B0 byte offset with no
		   manual arithmetic here) right after this object's own 7-slot IUnknown/
		   IDispatch bridge. */
		memcpy(
			(void*)((unsigned char*)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge)),
			pFormVtableTail,
			formIntrinsicPadding
		);
	}

	/* Copy the VB Specified vtable to the pWrapperVTable after bridgeStruct (and any
	   Form intrinsic padding) */
	memcpy(
		(void*)((unsigned int)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge) + formIntrinsicPadding),
		pvbNewData->opt.lpEvents,
		sizeof(void*) * uiVtableCount
	);

	/* Setup a vba_VBVTable struct, which is what we'll return and it'll be what VB code will use as 'this' (i.e. also contains the local storage of the class) */
	vba_VBVTable * ret = (vba_VBVTable*)malloc(sizeof(vba_VBVTable) + tObj->lpPublicBytes->iSize);
	if (ret == nullptr)
	{
		free(pWrapperVtable);
		return nullptr;
	}
	memset(
		ret,
		0,
		sizeof(vba_VBVTable) + tObj->lpPublicBytes->iSize
	);

	ret->lpVBVtable = pWrapperVtable;


	/* Create a vbClassWrapper (or, for a Form-derived class, the derived vbFormWrapper
	   -- see ClassWrapper.hpp/FormWrapper.hpp for why they're split), and assign the
	   vba_VBVTable pointer to it */
	ret->pWrapper = bIsFormLike
		? static_cast<vbClassWrapper*>(new vbFormWrapper(ret, pvbNewData))
		: new vbClassWrapper(ret, pvbNewData);

	DEBUG_WIDE(
		"ret %.8x, ret->lpVBVtable %.8x spans to %.8x, ret->pWrapper %.8x",
		(unsigned int)ret,
		(unsigned int)ret->lpVBVtable,
		(unsigned int)((unsigned int)ret->lpVBVtable + vtableBlobSize),
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
		"pvbNewData->opt.bWInitializeEvent %.8x, pvbNewData->opt.bWTerminateEvent %.8x, pvbNewData->opt.lpEvents %.8x, pvbNewData->opt.wEventCount %.8x",
		pvbNewData->opt.bWInitializeEvent,
		pvbNewData->opt.bWTerminateEvent,
		(unsigned int)(void*)pvbNewData->opt.lpEvents,
		pvbNewData->opt.wEventCount
	);

	ret->pWrapper->InvokeVB6Initialize();

	return ret;
} /* __vbaNew */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_AddRef(
	void			*arg1
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"arg1 %.8x",
		(unsigned int)arg1
	);

	return E_NOTIMPL;
} /* EVENT_SINK_AddRef */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_Release(
	void			*arg1
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"arg1 %.8x",
		(unsigned int)arg1
	);

	return E_NOTIMPL;
} /* EVENT_SINK_Release */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_QueryInterface(
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
} /* EVENT_SINK_QueryInterface */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_GetIDsOfNames(
	OLECHAR			*strIn,
	DWORD			unk1,
	DWORD			unk2,
	DWORD			unk3,
	DWORD			unk4,
	DWORD			unk5
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	return E_NOTIMPL;
} /* EVENT_SINK_GetIDsOfNames */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_Invoke(
	DWORD			*unk1,
	DWORD			unk2,
	DWORD			unk3,
	DWORD			unk4,
	DWORD			unk5,
	DWORD			unk6,
	DWORD			unk7,
	DWORD			unk8,
	DWORD			unk9
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	return E_NOTIMPL;
} /* EVENT_SINK_Invoke */

EXPORT HRESULT __stdcall Zombie_GetTypeInfo(
	DWORD			unk1,
	DWORD			unk2,
	DWORD			unk3,
	DWORD			unk4
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	return RPC_E_SERVER_DIED;
} /* Zombie_GetTypeInfo */

EXPORT HRESULT __stdcall Zombie_GetTypeInfoCount(
	DWORD			unk1,
	DWORD			unk2
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	return RPC_E_SERVER_DIED;
} /* Zombie_GetTypeInfoCount */

// https://bbs-vbstreets-ru.translate.goog/viewtopic.php?f=1&t=56212&start=0&hilit=GetMemObj&_x_tr_sch=http&_x_tr_sl=ru&_x_tr_tl=en&_x_tr_hl=en&_x_tr_pto=sc
//
// Real msvbvm60 disassembly (confirmed): GetMemEvent/PutMemEvent/SetMemEvent are the Get/Let/Set
// accessors for a WithEvents variable's own storage slot. The first argument is the OWNER class's
// own ObjectInfoWithOptional* (confirmed via runtime trace), the second is an (unused by our own
// implementation) index/flag, ppDst is the WithEvents field's own storage address, and pNewObj is
// the raw new object pointer (confirmed via runtime trace -- not a VARIANTARG, not a pointer-to-
// pointer). Real msvbvm60's SetMemEvent internally locates the connection via a late-bound,
// type-library-driven QueryInterface+Invoke dance this project doesn't have infrastructure for
// (see the plan); this implementation instead locates the compiler-built WithEvents sink block by
// signature-scanning the owner class's own compiled "Controls" data (EventDispatch.cpp) and wires
// it up through genuine IConnectionPointContainer/IConnectionPoint (confirmed real msvbvm60 also
// uses IConnectionPoint::Advise internally, at vtable offset 0x14).

/**
 * @brief			Returns the vbClassWrapper backing pObj, if pObj is genuinely one of our own
 *					wrapped objects (checked by confirming its vtable's first slot is our own
 *					BASIC_CLASS_QueryInterface bridge), nullptr otherwise (e.g. an external COM
 *					object, or garbage).
 */
static vbClassWrapper * TryGetWrapperOf(
	IDispatch		*pObj
)
{
	if (!pObj)
	{
		return nullptr;
	}

	vba_VBVTable *pAsVBVTable = (vba_VBVTable*)pObj;

	__try
	{
		void **pVtbl = (void**)pAsVBVTable->lpVBVtable;
		if (!pVtbl || pVtbl[0] != (void*)BASIC_CLASS_QueryInterface)
		{
			return nullptr;
		}
	}
	__except (EXCEPTION_EXECUTE_HANDLER)
	{
		return nullptr;
	}

	return pAsVBVTable->pWrapper;
}

HRESULT VBFormLoad(
	IDispatch		*object
)
{
	// dynamic_cast also rejects a genuine wrapper that just isn't a Form (a plain
	// class) -- Load on a non-Form object correctly fails here instead of silently
	// doing nothing, now that EnsureWindowCreated is Form-only.
	vbFormWrapper *pFormWrapper = dynamic_cast<vbFormWrapper*>(TryGetWrapperOf(object));
	if (!pFormWrapper)
	{
		return E_INVALIDARG;
	}

	return pFormWrapper->EnsureWindowCreated() ? S_OK : E_FAIL;
}

HRESULT VBFormUnload(
	IDispatch		*object
)
{
	vbFormWrapper *pFormWrapper = dynamic_cast<vbFormWrapper*>(TryGetWrapperOf(object));
	if (!pFormWrapper)
	{
		return E_INVALIDARG;
	}

	pFormWrapper->DestroyFormWindow();
	return S_OK;
}

bool VBFormTryQueryUnload(
	void	*pVBVTableRaw,
	short	*pCancel
)
{
	if (pCancel)
	{
		*pCancel = 0;
	}

	vba_VBVTable *pVBVTable = (vba_VBVTable*)pVBVTableRaw;
	if (!pVBVTable)
	{
		return false;
	}

	vbFormWrapper *pFormWrapper = dynamic_cast<vbFormWrapper*>(pVBVTable->pWrapper);
	if (!pFormWrapper)
	{
		return false;
	}

	return pFormWrapper->TryFireQueryUnload(pCancel);
}

/**
 * @brief			Disconnects the WithEvents variable's current value (if any) via a real
 *					IConnectionPoint::Unadvise, using the cookie msvbvm60 itself always stores at
 *					ppDst+4 (confirmed via disassembly of sub_66059B98).
 */
static void DisconnectWithEventsObject(
	IDispatch		*pOld,
	IDispatch		**ppDst
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	IConnectionPointContainer *pCPC = nullptr;
	if (SUCCEEDED(pOld->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC)) && pCPC)
	{
		IConnectionPoint *pCP = nullptr;
		if (SUCCEEDED(pCPC->FindConnectionPoint(IID_NULL, &pCP)) && pCP)
		{
			DWORD dwCookie = *((DWORD*)ppDst + 1);

			DEBUG_ASCII(
				"unadvising old sink %.8x, cookie %.8x",
				(unsigned int)pOld,
				dwCookie
			);

			pCP->Unadvise(dwCookie);
			pCP->Release();
		}
		pCPC->Release();
	}

	pOld->Release();
}

/**
 * @brief			Connects a new WithEvents value: stores it (addref'd) into *ppDst, then -- if a
 *					current owner instance is being tracked (see CurrentInstanceScope) and a
 *					compiler-built sink block can be found for this connection -- advises a new
 *					per-connection vbEventSinkInstance onto the new object's connection point, and
 *					stores the resulting cookie at ppDst+4 (matching real msvbvm60's own layout).
 */
static HRESULT ConnectWithEventsObject(
	ObjectInfoWithOptional	*pOwnerDescriptor,
	IDispatch				**ppDst,
	IDispatch				*pNewObj
)
{
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	IDispatch *pOld = *ppDst;
	if (pOld)
	{
		DisconnectWithEventsObject(pOld, ppDst);
	}

	*ppDst = pNewObj;

	if (!pNewObj)
	{
		return S_OK;
	}

	pNewObj->AddRef();

	void *pOwnerMe = GetCurrentInstance();
	if (!pOwnerMe)
	{
		DEBUG_ASCII(
			"no current owner instance tracked -- WithEvents connect outside Class_Initialize/"
			"Class_Terminate isn't supported yet; %.8x stored without connecting its events",
			(unsigned int)pNewObj
		);
		return S_OK;
	}

	vbClassWrapper *pNewWrapper = TryGetWrapperOf(pNewObj);
	if (!pNewWrapper)
	{
		DEBUG_ASCII(
			"pNewObj %.8x isn't one of our own wrapped objects; can't locate its sink block",
			(unsigned int)pNewObj
		);
		return S_OK;
	}

	IDispatch *pStaticSinkBlock = FindEventSinkBlock(pOwnerDescriptor, pNewWrapper->GetObjInfo());
	if (!pStaticSinkBlock)
	{
		DEBUG_ASCII("no compiler-built WithEvents sink block found for this connection");
		return S_OK;
	}

	IConnectionPointContainer *pCPC = nullptr;
	if (FAILED(pNewObj->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC)) || !pCPC)
	{
		DEBUG_ASCII("new source object doesn't support IConnectionPointContainer");
		return S_OK;
	}

	IConnectionPoint *pCP = nullptr;
	if (FAILED(pCPC->FindConnectionPoint(IID_NULL, &pCP)) || !pCP)
	{
		pCPC->Release();
		DEBUG_ASCII("new source object has no connection point");
		return S_OK;
	}

	vbEventSinkInstance *pSink = new vbEventSinkInstance(pStaticSinkBlock, pOwnerMe);
	DWORD dwCookie = 0;
	HRESULT hr = pCP->Advise(pSink, &dwCookie);

	DEBUG_ASCII(
		"Advise hr %.8x, sink %.8x, owner %.8x, cookie %.8x",
		(unsigned int)hr,
		(unsigned int)pSink,
		(unsigned int)pOwnerMe,
		dwCookie
	);

	if (SUCCEEDED(hr))
	{
		*((DWORD*)ppDst + 1) = dwCookie;
	}

	pSink->Release();
	pCP->Release();
	pCPC->Release();

	return S_OK;
}

/**
 * @brief			Reads the current value of a WithEvents variable's storage slot, addref'ing it
 *					for the caller.
 * @param			ppSrc			Pointer to the WithEvents variable's storage slot.
 * @param			ppDst			Receives an addref'd copy of *ppSrc.
 * @returns			S_OK.
 */
EXPORT HRESULT __stdcall GetMemEvent(
	DWORD			dwUnused1,
	DWORD			dwUnused2,
	IDispatch		**ppSrc,
	IDispatch		**ppDst
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ppSrc %.8x, ppDst %.8x",
		(unsigned int)ppSrc,
		(unsigned int)ppDst
	);

	IDispatch *pObj = *ppSrc;
	*ppDst = pObj;

	if (pObj)
	{
		pObj->AddRef();
	}

	return S_OK;
} /* GetMemEvent */

/**
 * @brief			Sets a WithEvents variable's storage slot to a new object.
 * @param			dwOwnerDescriptor	The owner class's own ObjectInfoWithOptional* (confirmed via
 *										runtime trace).
 * @param			ppDst			Pointer to the WithEvents variable's storage slot.
 * @param			pNewObj			The new object reference (confirmed via runtime trace: this is
 *									the raw object pointer itself, not a pointer to it).
 * @returns			S_OK.
 */
EXPORT HRESULT __stdcall PutMemEvent(
	DWORD			dwOwnerDescriptor,
	DWORD			dwUnused2,
	IDispatch		**ppDst,
	IDispatch		*pNewObj
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"dwOwnerDescriptor %.8x, dwUnused2 %.8x, ppDst %.8x, pNewObj %.8x",
		dwOwnerDescriptor,
		dwUnused2,
		(unsigned int)ppDst,
		(unsigned int)pNewObj
	);

	return ConnectWithEventsObject((ObjectInfoWithOptional*)dwOwnerDescriptor, ppDst, pNewObj);
} /* PutMemEvent */

/**
 * @brief			Sets a WithEvents variable's storage slot to a new object (Set-statement variant).
 * @param			dwOwnerDescriptor	The owner class's own ObjectInfoWithOptional* (confirmed via
 *										runtime trace).
 * @param			ppDst			Pointer to the WithEvents variable's storage slot.
 * @param			pNewObj			The new object reference (raw pointer, see PutMemEvent).
 * @returns			S_OK.
 */
EXPORT HRESULT __stdcall SetMemEvent(
	DWORD			dwOwnerDescriptor,
	DWORD			dwUnused2,
	IDispatch		**ppDst,
	IDispatch		*pNewObj
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"dwOwnerDescriptor %.8x, dwUnused2 %.8x, ppDst %.8x, pNewObj %.8x",
		dwOwnerDescriptor,
		dwUnused2,
		(unsigned int)ppDst,
		(unsigned int)pNewObj
	);

	return ConnectWithEventsObject((ObjectInfoWithOptional*)dwOwnerDescriptor, ppDst, pNewObj);
} /* SetMemEvent */

/**
 * @brief			Fires a class event: walks pMe's connection point sinks and calls Invoke(dispId,
 *					...) on each. ABI confirmed via disassembly of the real msvbvm60's
 *					__vbaRaiseEvent/RaiseEventOnBasicClass this session.
 * @param			unk1			pMe -- the raising object's own vba_VBVTable*.
 * @param			unk2			dispId of the event being raised (1-based, declaration order).
 * @param			argCount		Number of event arguments.
 * @param			...				argCount VARIANTARG values, passed by value.
 */
EXPORT HRESULT __vbaRaiseEvent(
	DWORD			unk1,
	DWORD			unk2,
	int				argCount,
	...
)
{
	va_list args;

	va_start(args, argCount);

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pMe %.8x, dispId %.8x, argCount = %d",
		unk1,
		unk2,
		argCount
	);

	vba_VBVTable *pMe = (vba_VBVTable*)unk1;
	if (pMe && pMe->pWrapper)
	{
		pMe->pWrapper->RaiseEvent((DISPID)unk2, (VARIANTARG*)args, (DWORD)argCount);
	}

	va_end(args);

	return S_OK;
} /* __vbaRaiseEvent */

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

	DEBUG_WIDE(
		"returning %.8x",
		(unsigned int)*ppiuDest
	);

	return *ppiuDest;
} /* __vbaObjSet */

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
} /* __vbaObjSetAddref */

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
		/* We don't have a IUnknown pointer, so return nothing? */
		return ppvObject;
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
} /* __vbaCastObj */

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

	if (!SUCCEEDED(hr))
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	IUnknown		*ppv;
	hr = CoCreateInstance(
		rCLSID,
		nullptr,
		CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
		IID_PPV_ARGS(&ppv)
	);
	DEBUG_WIDE(
		"CoCreateInstance = %.8x, ppv = %.8x",
		(unsigned int)hr,
		(unsigned int)ppv
	);

	if (!SUCCEEDED(hr))
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	IDispatch		*disp;
	hr = ppv->QueryInterface(
		IID_PPV_ARGS(&disp)
	);
	DEBUG_WIDE(
		"ppv->QueryInterface = %.8x, disp = %.8x",
		(unsigned int)hr,
		(unsigned int)disp
	);

	if (!SUCCEEDED(hr))
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	pvargObject->vt = VT_DISPATCH;
	pvargObject->pdispVal = disp;

	return hr;
} /* rtcCreateObject2 */

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
} /* __vbaObjVar */

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

	EXCEPINFO		excepInfo{};
	UINT			uArgErr = 0;

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

	if (!SUCCEEDED(hr))
	{
		DEBUG_WIDE(
			"pidObject->Invoke failed! GetLastError() = %.8x, excepInfo = '%ls', uArgErr = %.8x",
			GetLastError(),
			excepInfo.bstrDescription,
			uArgErr
		);

		vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR, &excepInfo); // TODO: Check if this exception is right
	}
} /* __vbaLateMemCall */

EXPORT void __stdcall __vbaVarLateMemSt(
	VARIANTARG		* pvargObject,
	BSTR			bstrMethodName,
	int				argCount,
	...
)
{
	HRESULT			hr;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"Object %.8x, bstrMethodName %.8x, argCount %.8x",
		(unsigned int)pvargObject,
		(unsigned int)bstrMethodName,
		argCount
	);


	IDispatch* pidObject;

	/* Get the IDispatch object from the Variant */
	pidObject = __vbaObjVar(pvargObject);

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

} /* __vbaVarLateMemSt */

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
} /* __vbaVarLateMemCallLdRf */

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
} /* objIDispatchGetDefaultValue */

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
} /* __vbaVarLateMemCallLd */

/**
 * @brief			Releases a COM Object (IUnknown) via its pointer, and nulls it.
 * @param			ppunkObj		Pointer to an object to release, and null the pointer.
 * @returns			Dereferenced object of ppunkObj.
 */
EXPORT void __fastcall __vbaFreeObj(
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
} /* __vbaFreeObj */

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
} /* __vbaFreeObjList */
