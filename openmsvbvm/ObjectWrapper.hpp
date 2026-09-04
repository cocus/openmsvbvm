#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "EventDispatch.hpp"

class vbObjectWrapper;

/**
 * The per-instance "vtable + bookkeeping" block __vbaNew (ObjectManipulation.cpp)
 * hands back as a VB6 object's own `this`. lpVBVtable is what compiled VB6 code (and
 * this project's own direct vtable-slot calls, e.g. Form.Show) actually dereferences
 * as the object's vtable; it must stay the first field, matching real COM/VB6 ABI.
 * lNull/pWrapper are this project's own bookkeeping, not part of the real msvbvm60
 * layout.
 */
typedef struct
{
    void* lpVBVtable;
    unsigned int lNull;
    class vbObjectWrapper* pWrapper;
} vba_VBVTable;

/**
 * Generic IDispatch/IConnectionPointContainer bridge for a compiled VB6 class
 * instance: everything a VB6 object needs regardless of whether it's a plain class,
 * a Form, or (eventually) any other kind: Class_Initialize/Class_Terminate, late-bound
 * GetIDsOfNames/Invoke, and the outgoing WithEvents connection point.
 *
 * Real msvbvm60 does NOT share an equivalent of this between a plain class and a
 * Form/control: `CClassModule` (plain classes) and `CTL` (the real base for
 * Form/MDIForm/every intrinsic control) are two completely independent hierarchies,
 * sharing no concrete base. `CClassModule`'s own constructor sets its vtables directly
 * with no base-class call at all, while `FORM` (a `CTL` subclass) calls `CTL::CTL(...)` first.
 *
 * This class is this project's own technical exception to that: `vbClassWrapper`
 * (plain classes, mirroring `CClassModule`) and `vbControlWrapper` (Forms/future
 * controls, mirroring `CTL`) both derive from `vbObjectWrapper` for real, DRY reasons
 * that don't apply to the real DLL. Neither of those two lineages is exposed via any
 * ABI a compiled VB6 EXE calls into directly (unlike, say, `_Form`'s vtable, which
 * genuinely has to match real byte offsets), so sharing one already-correct
 * implementation of the IDispatch/IConnectionPointContainer bridge here costs nothing
 * a real fidelity concern would care about, while duplicating ~150 lines of identical
 * logic (Class_Initialize/Terminate lookup, late-bound GetIDsOfNames/Invoke,
 * WithEvents' connection point) into two unrelated classes would.
 */
class vbObjectWrapper : IDispatch, IConnectionPointContainer
{
public:
    vbObjectWrapper(vba_VBVTable* pWrapperVtable, ObjectInfoWithOptional* pObjInfo);
    virtual ~vbObjectWrapper();

    // IUnknown interface
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppObj);

    ULONG __stdcall AddRef();
    ULONG __stdcall Release();

    // IDispatch interface
    HRESULT __stdcall GetTypeInfoCount(UINT* pctInfo);

    HRESULT __stdcall GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo);

    HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid);

    HRESULT __stdcall Invoke(
        DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);

    // IConnectionPointContainer interface
    HRESULT __stdcall EnumConnectionPoints(IEnumConnectionPoints** ppEnum);

    HRESULT __stdcall FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP);

    // VB6 Init and Terminate invokers
    void InvokeVB6Initialize();
    void InvokeVB6Terminate();

    // Fires dispId on every currently-advised sink of this object's connection point.
    void RaiseEvent(DISPID dispId, VARIANTARG* pArgs, DWORD argCount);

    ObjectInfoWithOptional* GetObjInfo() const
    {
        return m_pObjInfo;
    }

protected:
    // Accessible to derived wrappers (vbClassWrapper, vbControlWrapper, and anything
    // further derived from those, e.g. vbFormWrapper) that need the raw compiled
    // descriptor / owning vtable block; GetObjInfo() above covers unrelated callers.
    long m_nRefCount; // for managing the reference count
    int* m_pVtable = nullptr;
    vba_VBVTable* m_pVBVTable = nullptr;
    ObjectInfoWithOptional* m_pObjInfo = nullptr;
    vbConnectionPoint m_connectionPoint;
    bool m_bTerminating = false; // guards Release() against InvokeVB6Terminate() re-entering it
};
