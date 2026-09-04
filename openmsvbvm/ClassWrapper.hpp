#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "EventDispatch.hpp"

class vbClassWrapper;

/**
 * The per-instance "vtable + bookkeeping" block __vbaNew (ObjectManipulation.cpp)
 * hands back as a VB6 object's own `this`. lpVBVtable is what compiled VB6 code (and
 * this project's own direct vtable-slot calls, e.g. Form.Show) actually dereferences
 * as the object's vtable -- it must stay the first field, matching real COM/VB6 ABI.
 * lNull/pWrapper are this project's own bookkeeping, not part of the real msvbvm60
 * layout.
 */
typedef struct
{
	void					* lpVBVtable;
	unsigned int			lNull;
	class vbClassWrapper	* pWrapper;
} vba_VBVTable;

/**
 * Generic IDispatch/IConnectionPointContainer bridge for a compiled VB6 class
 * instance -- everything a VB6 object needs regardless of whether it's a plain class,
 * a Form, or (eventually) any other kind: Class_Initialize/Class_Terminate, late-bound
 * GetIDsOfNames/Invoke, and the outgoing WithEvents connection point.
 *
 * Form-specific behavior (a real Win32 window, Load/Show/QueryUnload dispatch) used to
 * live directly in this class, gated by an bIsFormLike flag -- that's been split out
 * into the derived vbFormWrapper (FormWrapper.hpp) instead, so this class stays
 * meaningful for every VB6 object, not just ones that happen to also be Forms.
 * __vbaNew (ObjectManipulation.cpp) picks which one to construct based on
 * IsFormLikeDescriptor.
 */
class vbClassWrapper : IDispatch, IConnectionPointContainer
{
public:
	vbClassWrapper(vba_VBVTable * pWrapperVtable, ObjectInfoWithOptional* pObjInfo);
	virtual ~vbClassWrapper();

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

	// IConnectionPointContainer interface
	HRESULT __stdcall EnumConnectionPoints(
		IEnumConnectionPoints ** ppEnum
	);

	HRESULT __stdcall FindConnectionPoint(
		REFIID riid,
		IConnectionPoint ** ppCP
	);

	// VB6 Init and Terminate invokers
	void InvokeVB6Initialize();
	void InvokeVB6Terminate();

	// Fires dispId on every currently-advised sink of this object's connection point.
	void RaiseEvent(DISPID dispId, VARIANTARG * pArgs, DWORD argCount);

	ObjectInfoWithOptional * GetObjInfo() const { return m_pObjInfo; }

protected:
	// Accessible to derived wrappers (e.g. vbFormWrapper) that need the raw compiled
	// descriptor / owning vtable block; GetObjInfo() above covers unrelated callers.
	long						m_nRefCount;   // for managing the reference count
	int							* m_pVtable = nullptr;
	vba_VBVTable				* m_pVBVTable = nullptr;
	ObjectInfoWithOptional*		m_pObjInfo = nullptr;
	vbConnectionPoint			m_connectionPoint;
};
