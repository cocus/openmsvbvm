#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"
#include <wctype.h>

#include "vba_Locale.h"
#include "ObjectManipulation.hpp"
#include "DllObjectInterface.hpp"

#include "vba_structures.h"
#include "EventDispatch.hpp"
#include "FormWindow.hpp"
#include "FormProperties.hpp"
#include "ObjectWrapper.hpp"
#include "ClassWrapper.hpp"
#include "ControlWrapper.hpp"
#include "FormWrapper.hpp"

#define DECLARE_BASIC_CLASS_WRAPPER(name, arg_types, arg_names, ret_type) \
    EXPORT ret_type __stdcall BASIC_CLASS_##name(arg_types) \
    { \
\
        if (pFakeVtable->pWrapper) \
        { \
            LOG(LOG_TRACE) << L"pFakeVtable " << vbl::Hex(pFakeVtable) << L", pWrapper " << vbl::Hex(pFakeVtable->pWrapper); \
            return pFakeVtable->pWrapper->##name(arg_names); \
        } \
        else \
        { \
            LOG(LOG_TRACE) << L"pFakeVtable " << vbl::Hex(pFakeVtable); \
        } \
\
        return E_NOTIMPL; \
    }

#define DECLARE_BASIC_CLASS_TYPES_FOR_STRUCT(name, arg_types, arg_names, ret_type) ret_type(__stdcall* p##name)(arg_types);

#define ASSIGN_MEMBERS_OF_BRIDGE_STRUCT(name, arg_types, arg_names, ret_type) bridgeStruct.p##name = BASIC_CLASS_##name;

#define BASIC_CLASS_WRAPPER_FUNCTIONS(decl) \
    decl(QueryInterface, ARGS(vba_VBVTable* pFakeVtable, REFIID riid, void** ppObj), ARGS(riid, ppObj), ULONG) \
        decl(AddRef, ARGS(vba_VBVTable* pFakeVtable), ARGS(), ULONG) decl(Release, ARGS(vba_VBVTable* pFakeVtable), ARGS(), ULONG) \
            decl(GetTypeInfoCount, ARGS(vba_VBVTable* pFakeVtable, UINT* pctInfo), ARGS(pctInfo), HRESULT) decl( \
                GetTypeInfo, ARGS(vba_VBVTable* pFakeVtable, UINT itinfo, LCID lcid, ITypeInfo** pptinfo), ARGS(itinfo, lcid, pptinfo), HRESULT) \
                decl( \
                    GetIDsOfNames, \
                    ARGS(vba_VBVTable* pFakeVtable, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid), \
                    ARGS(riid, rgszNames, cNames, lcid, rgdispid), \
                    HRESULT) \
                    decl( \
                        Invoke, \
                        ARGS( \
                            vba_VBVTable* pFakeVtable, \
                            DISPID dispidMember, \
                            REFIID riid, \
                            LCID lcid, \
                            WORD wFlags, \
                            DISPPARAMS* pdispparams, \
                            VARIANT* pvarResult, \
                            EXCEPINFO* pexcepinfo, \
                            UINT* puArgErr), \
                        ARGS(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr), \
                        HRESULT)

/* This spawns the exported bridge functions to the BASIC_CLASS wrapper */
BASIC_CLASS_WRAPPER_FUNCTIONS(DECLARE_BASIC_CLASS_WRAPPER);

/* This structure holds pointer to all the BASIC_CLASS wrapper functions */
typedef struct
{
    BASIC_CLASS_WRAPPER_FUNCTIONS(DECLARE_BASIC_CLASS_TYPES_FOR_STRUCT);
} vba_BASIC_CLASS_IUnknownBridge;

ULONG __stdcall vbObjectWrapper::AddRef()
{
    LOG_OBJ(LOG_DEBUG, this) << L"m_nRefCount was " << vbl::Hex((unsigned long)m_nRefCount);

    return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall vbObjectWrapper::Release()
{
    LOG_OBJ(LOG_DEBUG, this) << L"m_nRefCount was " << vbl::Hex((unsigned long)m_nRefCount);

    long nRefCount = 0;
    nRefCount = InterlockedDecrement(&m_nRefCount);
    if (nRefCount == 0 && m_bTerminating == false)
    {
        m_bTerminating = true;

        InvokeVB6Terminate(); // This will re-call Release, but we already set m_bTerminating

        delete this;
    }

    return nRefCount;
}

void vbObjectWrapper::InvokeVB6Initialize()
{

    typedef void(__stdcall * pInit)(vba_VBVTable * ths);

    /* bWInitializeEvent/bWTerminateEvent are byte offsets measured from the wEventCount field
       itself (8 bytes before lpEvents), NOT plain indices into lpEvents and NOT always the last
       two entries -- a class with WithEvents members appends its Get/Put/Set accessor thunks
       after Class_Initialize/Class_Terminate, so "last two slots" only works by coincidence for
       classes without WithEvents. Confirmed via msvbvm60 disassembly + a real WithEvents test case. */
    unsigned int index = ((unsigned int)this->m_pObjInfo->opt.bWInitializeEvent - 8) / sizeof(void*);

    // Confirmed live: this byte-offset formula only resolves to a real slot inside
    // lpEvents for a plain class -- for a Form (e.g. a Form with no WithEvents members
    // and a tiny lpEvents array), the same bWInitializeEvent/bWTerminateEvent values
    // compute an index far past wEventCount, landing on unrelated compiled data
    // (confirmed via raw disassembly: not a real thunk) and crashing with an illegal
    // instruction when called. Bound the index against wEventCount before trusting it.
    if (index >= this->m_pObjInfo->opt.wEventCount)
    {
        return;
    }

    unsigned int** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->opt.lpEvents) + index;

    pInit pCall = (pInit)(*pAddr);

    LOG(LOG_DEBUG) << L"pCall " << vbl::Hex((unsigned long)pCall);

    // Not every class implements Class_Initialize -- lpEvents can hold a null slot.
    // Calling through that is an unconditional crash, so skip when there's nothing to call.
    if (pCall == nullptr)
    {
        return;
    }

    CurrentInstanceScope instanceScope(this->m_pVBVTable);
    pCall(this->m_pVBVTable);
}

void vbObjectWrapper::InvokeVB6Terminate()
{

    typedef void(__stdcall * pInit)(vba_VBVTable * ths);

    unsigned int index = ((unsigned int)this->m_pObjInfo->opt.bWTerminateEvent - 8) / sizeof(void*);

    // See the matching guard + comment in InvokeVB6Initialize.
    if (index >= this->m_pObjInfo->opt.wEventCount)
    {
        return;
    }

    unsigned int** pAddr = (unsigned int**)(void*)(this->m_pObjInfo->opt.lpEvents) + index;

    pInit pCall = (pInit)(*pAddr);

    LOG(LOG_DEBUG) << L"pCall " << vbl::Hex((unsigned long)pCall);

    // Not every class implements Class_Terminate, so a null slot here is expected
    // and must not be called.
    if (pCall == nullptr)
    {
        return;
    }

    CurrentInstanceScope instanceScope(this->m_pVBVTable);
    pCall(this->m_pVBVTable);
}

vbObjectWrapper::vbObjectWrapper(vba_VBVTable* pWrapperVtable, ObjectInfoWithOptional* pObjInfo) :
    m_nRefCount(1), m_pVBVTable(pWrapperVtable), m_pObjInfo(pObjInfo), m_connectionPoint(static_cast<IDispatch*>(this))
{

    LOG_OBJ(LOG_DEBUG, this) << L"New wrapper object! pWrapperVtable " << vbl::Hex((unsigned long)pWrapperVtable)
                             << L", pObjInfo " << vbl::Hex((unsigned long)pObjInfo);
}

vbObjectWrapper::~vbObjectWrapper()
{

    LOG_OBJ(LOG_DEBUG, this) << L"vbObjectWrapper::~vbObjectWrapper()!";

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

HRESULT __stdcall vbObjectWrapper::QueryInterface(REFIID riid, void** ppObj)
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

HRESULT __stdcall vbObjectWrapper::GetTypeInfoCount(UINT* pctInfo)
{
    if (pctInfo)
    {
        *pctInfo = 0; /* no real ITypeInfo object provided yet */
    }
    return S_OK;
}

HRESULT __stdcall vbObjectWrapper::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
{
    return TYPE_E_ELEMENTNOTFOUND;
}

HRESULT __stdcall vbObjectWrapper::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid)
{

    if (!m_pObjInfo || !m_pObjInfo->hdr.lpObject || !rgszNames || !rgdispid)
    {
        return E_NOTIMPL;
    }

    /* Real name -> DISPID resolution sourced from this compiled class's own
       PublicObjectDescriptor.lpMethodNames table (already-populated for native
       builds, including for Private WithEvents handler subs. */
    PublicObjectDescriptor* pDesc = m_pObjInfo->hdr.lpObject;
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

    LOG(LOG_DEBUG) << L"cNames " << vbl::Hex((unsigned long)cNames) << L", hr " << vbl::Hres(hrOverall);

    return hrOverall;
}

HRESULT __stdcall vbObjectWrapper::Invoke(
    DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
    return E_NOTIMPL;
}

HRESULT __stdcall vbObjectWrapper::EnumConnectionPoints(IEnumConnectionPoints** ppEnum)
{
    return E_NOTIMPL;
}

HRESULT __stdcall vbObjectWrapper::FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP)
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

void vbObjectWrapper::RaiseEvent(DISPID dispId, VARIANTARG* pArgs, DWORD argCount)
{
    RaiseEventOnSinks(m_connectionPoint.Sinks(), dispId, pArgs, argCount);
}

/**
 * @brief           TBD
 * @param           TBD
 * @returns         TBD
 */
EXPORT void __stdcall __vbaHresultCheckObj(int arg1, int arg2, GUID* arg3, __int16 arg4)
{

    LOG(LOG_DEBUG) << L"arg1 " << vbl::Hex((unsigned long)arg1) << L", arg2 " << vbl::Hex((unsigned long)arg2) << L", arg3 "
                   << vbl::Hex((unsigned long)arg3) << L", arg4 " << vbl::Hex((unsigned long)arg4);

    // arg1 is the HRESULT returned by the object call this guards (e.g. Load/Unload
    // on the VB global object). Raise the matching VB runtime error instead, same
    // as every other HRESULT-checking call site in this project.
    HRESULT hr = (HRESULT)arg1;
    if (FAILED(hr))
    {
        vbaRaiseException(vbaErrorFromHRESULT(hr));
    }
} /* __vbaHresultCheckObj */

/**
 * @brief           Returns true if pDesc describes a Form- or MDIForm-derived class.
 *
 * PublicObjectDescriptor.fObjectType (see that field's own doc comment in
 * vba_structures.h for the full confirmed bit table) is the real, stable signal: bit
 * 0x80 is the discriminator, confirmed shared by both a plain Form and an MDIForm.
 */
static bool IsFormLikeDescriptor(ObjectInfoWithOptional* pDesc)
{
    if (!pDesc->hdr.lpObject)
    {
        return false;
    }

    return (pDesc->hdr.lpObject->fObjectType & 0x80) != 0;
}

/**
 * A Form's own compiled code reaches a placed control by NAME (e.g. "Label1.Caption
 * = ...") through a direct vtable call, exactly like ".Show" -- confirmed live via
 * disassembly of "Label1.Caption = Now": it compiles to `call dword ptr [Me+0x2FCh]`
 * (a project with a single referenced control) or `[Me+0x300h]` (a second project
 * with an earlier-declared, unreferenced Command1 ahead of it) -- i.e. ONE reserved
 * slot per REAL placed control (see CountPlacedControls/FindPlacedControlByAccessorIndex,
 * EventDispatch.hpp), in opt.lpControls's own declaration order, starting right
 * after _Form's own confirmed tail PLUS exactly one more slot whose real meaning
 * isn't confirmed (byte 0x2F8 = _Form's own last slot's end; the first control
 * accessor lands at 0x2FC, not 0x2F8 -- that one extra slot is left null/unused
 * below rather than guessed at).
 */
template <int N>
static IDispatch* __stdcall BASIC_CLASS_GetControlAccessor(vba_VBVTable* pFakeVtable)
{
    if (!pFakeVtable || !pFakeVtable->pWrapper)
    {
        return nullptr;
    }

    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(pFakeVtable->pWrapper);
    if (!pFormWrapper)
    {
        return nullptr;
    }

    return pFormWrapper->GetControlAccessor(N);
}

/* A fixed-size table of the above, one per supported accessor index -- C++ has no
   way to generate an arbitrary number of distinct functions at compile time, so this
   is also this project's practical limit on placed controls actually reachable by
   name from compiled code. 32 comfortably covers any test project built so far. */
#define CONTROL_ACCESSOR_TRAMPOLINE_LIST(X) \
    X(0) \
    X(1) \
    X(2) \
    X(3) \
    X(4) \
    X(5) \
    X(6) \
    X(7) \
    X(8) \
    X(9) \
    X(10) \
    X(11) X(12) X(13) X(14) X(15) X(16) X(17) X(18) X(19) X(20) X(21) X(22) X(23) X(24) X(25) X(26) X(27) X(28) X(29) X(30) X(31)
#define CONTROL_ACCESSOR_ARRAY_ENTRY(N) &BASIC_CLASS_GetControlAccessor<N>,

static IDispatch*(__stdcall* const g_controlAccessorTrampolines[])(vba_VBVTable*) = {
    CONTROL_ACCESSOR_TRAMPOLINE_LIST(CONTROL_ACCESSOR_ARRAY_ENTRY)};

static const size_t kMaxControlAccessorTrampolines = sizeof(g_controlAccessorTrampolines) / sizeof(g_controlAccessorTrampolines[0]);

/**
 * @brief           Constructs a wrapper object for a VB6 object, and instantiates it.
 * @param           pvbNewData          Object info pointer.
 * @returns         Valid vba_VBVTable pointer on success, nullptr otherwise.
 */
EXPORT vba_VBVTable* __stdcall __vbaNew(ObjectInfoWithOptional* pvbNewData)
{

    if (pvbNewData == nullptr)
    {
        LOG(LOG_WARN) << L"pvbNewData " << vbl::Hex(pvbNewData);

        return nullptr;
    }

    LOG(LOG_DEBUG) << L"pvbNewData " << vbl::Hex(pvbNewData) << L", pvbNewData->iEventCount "
                   << vbl::Hex(pvbNewData->opt.wEventCount) << L", pvbNewData->hdr.lpObject " << vbl::Hex(pvbNewData->hdr.lpObject)
                   << L", pvbNewData->opt.lpBasicClassObject " << vbl::Hex(pvbNewData->opt.lpBasicClassObject);

    if (!pvbNewData->hdr.lpObject)
    {
        return nullptr;
    }

    PublicObjectDescriptor* tObj = static_cast<PublicObjectDescriptor*>(pvbNewData->hdr.lpObject);
    LOG(LOG_DEBUG) << L"tObj->lpModulePublic " << vbl::Hex((unsigned long)tObj->lpModulePublic) << L", tObj->lpModuleStatic "
                   << vbl::Hex((unsigned long)tObj->lpModuleStatic) << L", tObj->lpPublicBytes "
                   << vbl::Hex((unsigned long)tObj->lpPublicBytes) << L", tObj->lpStaticBytes "
                   << vbl::Hex((unsigned long)tObj->lpStaticBytes) << L", tObj->dwMethodCount "
                   << vbl::Hex((unsigned long)tObj->dwMethodCount) << L", tObj->bStaticVars "
                   << vbl::Hex((unsigned long)tObj->bStaticVars);

    if (tObj->lpPublicBytes == nullptr)
    {
        return nullptr;
    }

    LOG(LOG_DEBUG) << L"tObj->lpPublicBytes->iConst1 " << vbl::Hex((unsigned long)tObj->lpPublicBytes->iConst1)
                   << L", tObj->lpPublicBytes->iSize " << vbl::Hex((unsigned long)tObj->lpPublicBytes->iSize);

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
    void* const* pFormVtableTail = nullptr;
    size_t formIntrinsicPadding = 0;
    if (bIsFormLike)
    {
        pFormVtableTail = GetFormInterfaceVtableTail(&formVtableTailSlotCount);
        formIntrinsicPadding = formVtableTailSlotCount * sizeof(void*);
        vtableBlobSize += formIntrinsicPadding;
    }

    /* Right after _Form's own tail: one unexplained-but-confirmed-necessary null
       slot, then one real "get this placed control by name" accessor slot per
       control (see BASIC_CLASS_GetControlAccessor's own doc comment above for why --
       this used to alias pvbNewData->opt.lpEvents below instead, which is the actual
       bug that made a control-referencing handler recurse into itself). Only
       Form-like objects have placed controls at all. */
    size_t numPlacedControls = 0;
    size_t controlAccessorPadding = 0;
    if (bIsFormLike)
    {
        numPlacedControls = CountPlacedControls(pvbNewData);
        if (numPlacedControls > kMaxControlAccessorTrampolines)
        {
            LOG(LOG_WARN) << L"Too many placed controls (" << numPlacedControls << L") -- only " << kMaxControlAccessorTrampolines
                          << L" are supported, so truncating to that limit.";
            numPlacedControls = kMaxControlAccessorTrampolines;
        }
        controlAccessorPadding = (1 + numPlacedControls) * sizeof(void*);
        vtableBlobSize += controlAccessorPadding;
    }

    /* pWrapperVTable is technically the "VB class" with its functions after the IDispatch stuff */
    void* pWrapperVtable = malloc(vtableBlobSize);
    if (pWrapperVtable == nullptr)
    {
        return nullptr;
    }
    memset(pWrapperVtable, 0, vtableBlobSize);

    /* Setup the bridgeStruct (IUnk-like that bridges VB's VTable to COM) */
    vba_BASIC_CLASS_IUnknownBridge bridgeStruct;
    BASIC_CLASS_WRAPPER_FUNCTIONS(ASSIGN_MEMBERS_OF_BRIDGE_STRUCT);

    /* And copy it to the pWrapperVTable */
    memcpy(pWrapperVtable, &bridgeStruct, sizeof(vba_BASIC_CLASS_IUnknownBridge));

    if (bIsFormLike && pFormVtableTail != nullptr)
    {
        /* Landing every real _Form member right after this object's own 7-slot IUnknown/IDispatch bridge. */
        memcpy((void*)((unsigned char*)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge)), pFormVtableTail, formIntrinsicPadding);
    }

    if (bIsFormLike && numPlacedControls > 0)
    {
        /* One trampoline per placed control, starting one slot after _Form's own
           tail (that first slot stays null) */
        void** pAccessorSlots =
            (void**)((unsigned char*)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge) + formIntrinsicPadding + sizeof(void*));

        for (size_t i = 0; i < numPlacedControls; i++)
        {
            pAccessorSlots[i] = (void*)g_controlAccessorTrampolines[i];
        }
    }

    /* Copy the VB Specified vtable to the pWrapperVTable after bridgeStruct (and any
       Form intrinsic/control-accessor padding) */
    memcpy(
        (void*)((unsigned int)pWrapperVtable + sizeof(vba_BASIC_CLASS_IUnknownBridge) + formIntrinsicPadding + controlAccessorPadding),
        pvbNewData->opt.lpEvents,
        sizeof(void*) * uiVtableCount);

    /* Setup a vba_VBVTable struct, which is what we'll return and it'll be what VB code will use as 'this' (i.e. also contains the local storage of the class) */
    vba_VBVTable* ret = (vba_VBVTable*)malloc(sizeof(vba_VBVTable) + tObj->lpPublicBytes->iSize);
    if (ret == nullptr)
    {
        free(pWrapperVtable);
        return nullptr;
    }
    memset(ret, 0, sizeof(vba_VBVTable) + tObj->lpPublicBytes->iSize);

    ret->lpVBVtable = pWrapperVtable;

    /* Create a vbClassWrapper (plain class, ClassWrapper.hpp which mirrors real
       msvbvm60's CClassModule) or, for a Form-derived class, a vbFormWrapper
       (FormWrapper.hpp, itself a vbControlWrapper which mirrors real msvbvm60's FORM
       deriving from CTL). */
    ret->pWrapper = bIsFormLike ? static_cast<vbObjectWrapper*>(new vbFormWrapper(ret, pvbNewData))
                                : static_cast<vbObjectWrapper*>(new vbClassWrapper(ret, pvbNewData));

    LOG(LOG_DEBUG) << L"ret " << vbl::Hex(ret) << L", ret->lpVBVtable " << vbl::Hex(ret->lpVBVtable) << L", vtableBlobSize "
                   << vtableBlobSize << L", ret->pWrapper " << vbl::Hex(ret->pWrapper);

    if (ret->pWrapper == nullptr)
    {
        free(pWrapperVtable);
        free(ret);
        return nullptr;
    }

    /* Call the Initialize method */
    LOG(LOG_DEBUG) << L"pvbNewData->opt.bWInitializeEvent " << vbl::Hex(pvbNewData->opt.bWInitializeEvent)
                   << L", pvbNewData->opt.bWTerminateEvent " << vbl::Hex(pvbNewData->opt.bWTerminateEvent)
                   << L", pvbNewData->opt.lpEvents " << vbl::Hex(pvbNewData->opt.lpEvents) << L", pvbNewData->opt.wEventCount "
                   << vbl::Hex(pvbNewData->opt.wEventCount);

    ret->pWrapper->InvokeVB6Initialize();

    return ret;
} /* __vbaNew */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_AddRef(void* arg1)
{
    LOG(LOG_ERROR) << L"arg1 " << vbl::Hex(arg1);

    return E_NOTIMPL;
} /* EVENT_SINK_AddRef */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_Release(void* arg1)
{
    LOG(LOG_ERROR) << L"arg1 " << vbl::Hex(arg1);

    return E_NOTIMPL;
} /* EVENT_SINK_Release */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_QueryInterface(void* arg1, REFIID arg2, void** ppvObject)
{
    LOG(LOG_ERROR) << L"arg1 " << vbl::Hex(arg1) << L", arg2 " << vbl::Hex(&arg2) << L", ppvObject " << vbl::Hex(ppvObject);

    return E_NOTIMPL;
} /* EVENT_SINK_QueryInterface */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_GetIDsOfNames(OLECHAR* strIn, DWORD unk1, DWORD unk2, DWORD unk3, DWORD unk4, DWORD unk5)
{
    LOG(LOG_ERROR) << L"?";

    return E_NOTIMPL;
} /* EVENT_SINK_GetIDsOfNames */

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall EVENT_SINK_Invoke(DWORD* unk1, DWORD unk2, DWORD unk3, DWORD unk4, DWORD unk5, DWORD unk6, DWORD unk7, DWORD unk8, DWORD unk9)
{
    LOG(LOG_ERROR) << L"?";

    return E_NOTIMPL;
} /* EVENT_SINK_Invoke */

EXPORT HRESULT __stdcall Zombie_GetTypeInfo(DWORD unk1, DWORD unk2, DWORD unk3, DWORD unk4)
{
    LOG(LOG_ERROR) << L"?";

    return RPC_E_SERVER_DIED;
} /* Zombie_GetTypeInfo */

EXPORT HRESULT __stdcall Zombie_GetTypeInfoCount(DWORD unk1, DWORD unk2)
{
    LOG(LOG_ERROR) << L"?";

    return RPC_E_SERVER_DIED;
} /* Zombie_GetTypeInfoCount */

/* https://bbs-vbstreets-ru.translate.goog/viewtopic.php?f=1&t=56212&start=0&hilit=GetMemObj&_x_tr_sch=http&_x_tr_sl=ru&_x_tr_tl=en&_x_tr_hl=en&_x_tr_pto=sc

 GetMemEvent/PutMemEvent/SetMemEvent are the Get/Let/Set accessors for a WithEvents variable's
 own storage slot. The first argument is the OWNER class's own ObjectInfoWithOptional*, the second
 is an index/flag, ppDst is the WithEvents field's own storage address, and pNewObj is
 the raw new object pointer. Real msvbvm60's SetMemEvent internally locates the connection via
 a late-bound, type-library-driven QueryInterface+Invoke dance this project doesn't have
 infrastructure for; this implementation instead locates the compiler-built WithEvents sink block by
 signature-scanning the owner class's own compiled "Controls" data (EventDispatch.cpp) and wires
 it up through genuine IConnectionPointContainer/IConnectionPoint, which real msvbvm60 also
 uses IConnectionPoint::Advise internally, at vtable offset 0x14. */

/**
 * @brief			Returns the vbObjectWrapper backing pObj, if pObj is genuinely one
 *					of our own wrapped objects (checked by confirming its vtable's
 *					first slot is our own BASIC_CLASS_QueryInterface bridge), nullptr
 *					otherwise (e.g. an external COM object, or garbage). The returned
 *					pointer's actual dynamic type is vbClassWrapper or vbFormWrapper
 *					(or, in the future, some other vbControlWrapper-derived type) --
 *					dynamic_cast to whichever one a caller actually needs.
 */
static vbObjectWrapper* TryGetWrapperOf(IDispatch* pObj)
{
    if (!pObj)
    {
        return nullptr;
    }

    vba_VBVTable* pAsVBVTable = (vba_VBVTable*)pObj;

    __try
    {
        void** pVtbl = (void**)pAsVBVTable->lpVBVtable;
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

HRESULT VBFormLoad(IDispatch* object)
{
    // dynamic_cast also rejects a genuine wrapper that just isn't a Form (a plain
    // class) -- Load on a non-Form object correctly fails here instead of silently
    // doing nothing, now that EnsureWindowCreated is Form-only.
    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(TryGetWrapperOf(object));
    if (!pFormWrapper)
    {
        return E_INVALIDARG;
    }

    return pFormWrapper->EnsureWindowCreated() ? S_OK : E_FAIL;
}

HRESULT VBFormUnload(IDispatch* object)
{
    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(TryGetWrapperOf(object));
    if (!pFormWrapper)
    {
        return E_INVALIDARG;
    }

    pFormWrapper->DestroyFormWindow();
    return S_OK;
}

HWND VBFormGetHwnd(IDispatch* object)
{
    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(TryGetWrapperOf(object));
    if (!pFormWrapper)
    {
        return nullptr;
    }

    return pFormWrapper->GetHwnd();
}

bool VBFormTryQueryUnload(void* pVBVTableRaw, short* pCancel)
{
    if (pCancel)
    {
        *pCancel = 0;
    }

    vba_VBVTable* pVBVTable = (vba_VBVTable*)pVBVTableRaw;
    if (!pVBVTable)
    {
        return false;
    }

    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(pVBVTable->pWrapper);
    if (!pFormWrapper)
    {
        return false;
    }

    return pFormWrapper->TryFireQueryUnload(pCancel);
}

void VBFormHandleCommand(void* pVBVTableRaw, WORD controlId, WORD notifyCode)
{
    vba_VBVTable* pVBVTable = (vba_VBVTable*)pVBVTableRaw;
    if (!pVBVTable)
    {
        return;
    }

    vbFormWrapper* pFormWrapper = dynamic_cast<vbFormWrapper*>(pVBVTable->pWrapper);
    if (!pFormWrapper)
    {
        return;
    }

    pFormWrapper->HandleCommand(controlId, notifyCode);
}

/**
 * @brief			Disconnects the WithEvents variable's current value (if any) via a real
 *					IConnectionPoint::Unadvise, using the cookie msvbvm60 itself always stores at
 *					ppDst+4 (confirmed via disassembly of sub_66059B98).
 */
static void DisconnectWithEventsObject(IDispatch* pOld, IDispatch** ppDst)
{

    IConnectionPointContainer* pCPC = nullptr;
    if (SUCCEEDED(pOld->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC)) && pCPC)
    {
        IConnectionPoint* pCP = nullptr;
        if (SUCCEEDED(pCPC->FindConnectionPoint(IID_NULL, &pCP)) && pCP)
        {
            DWORD dwCookie = *((DWORD*)ppDst + 1);

            LOG(LOG_DEBUG) << L"unadvising old sink " << vbl::Hex((unsigned long)pOld) << L", cookie "
                           << vbl::Hex((unsigned long)dwCookie);

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
static HRESULT ConnectWithEventsObject(ObjectInfoWithOptional* pOwnerDescriptor, IDispatch** ppDst, IDispatch* pNewObj)
{

    IDispatch* pOld = *ppDst;
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

    void* pOwnerMe = GetCurrentInstance();
    if (!pOwnerMe)
    {
        LOG(LOG_WARN) << L"no current owner instance tracked, WithEvents connect outside Class_Initialize/"
                         L"Class_Terminate isn't supported yet; "
                      << vbl::Hex((unsigned long)pNewObj) << L" stored without connecting its events";
        return S_OK;
    }

    vbObjectWrapper* pNewWrapper = TryGetWrapperOf(pNewObj);
    if (!pNewWrapper)
    {
        LOG(LOG_WARN) << L"pNewObj " << vbl::Hex((unsigned long)pNewObj)
                      << L" isn't one of our own wrapped objects; can't locate its sink block";
        return S_OK;
    }

    IDispatch* pStaticSinkBlock = FindEventSinkBlock(pOwnerDescriptor, pNewWrapper->GetObjInfo());
    if (!pStaticSinkBlock)
    {
        LOG(LOG_WARN) << L"no compiler-built WithEvents sink block found for this connection";
        return S_OK;
    }

    IConnectionPointContainer* pCPC = nullptr;
    if (FAILED(pNewObj->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC)) || !pCPC)
    {
        LOG(LOG_DEBUG) << L"new source object doesn't support IConnectionPointContainer";
        return S_OK;
    }

    IConnectionPoint* pCP = nullptr;
    if (FAILED(pCPC->FindConnectionPoint(IID_NULL, &pCP)) || !pCP)
    {
        pCPC->Release();
        LOG(LOG_DEBUG) << L"new source object has no connection point";
        return S_OK;
    }

    vbEventSinkInstance* pSink = new vbEventSinkInstance(pStaticSinkBlock, pOwnerMe);
    DWORD dwCookie = 0;
    HRESULT hr = pCP->Advise(pSink, &dwCookie);

    LOG(LOG_DEBUG) << L"Advise hr " << vbl::Hres(hr) << L", sink " << vbl::Hex((unsigned long)pSink) << L", owner "
                   << vbl::Hex((unsigned long)pOwnerMe) << L", cookie " << vbl::Hex((unsigned long)dwCookie);

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
EXPORT HRESULT __stdcall GetMemEvent(DWORD dwUnused1, DWORD dwUnused2, IDispatch** ppSrc, IDispatch** ppDst)
{

    LOG(LOG_DEBUG) << L"ppSrc " << vbl::Hex((unsigned long)ppSrc) << L", ppDst " << vbl::Hex((unsigned long)ppDst);

    IDispatch* pObj = *ppSrc;
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
EXPORT HRESULT __stdcall PutMemEvent(DWORD dwOwnerDescriptor, DWORD dwUnused2, IDispatch** ppDst, IDispatch* pNewObj)
{

    LOG(LOG_DEBUG) << L"dwOwnerDescriptor " << vbl::Hex((unsigned long)dwOwnerDescriptor) << L", dwUnused2 "
                   << vbl::Hex((unsigned long)dwUnused2) << L", ppDst " << vbl::Hex((unsigned long)ppDst) << L", pNewObj "
                   << vbl::Hex((unsigned long)pNewObj);

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
EXPORT HRESULT __stdcall SetMemEvent(DWORD dwOwnerDescriptor, DWORD dwUnused2, IDispatch** ppDst, IDispatch* pNewObj)
{

    LOG(LOG_DEBUG) << L"dwOwnerDescriptor " << vbl::Hex((unsigned long)dwOwnerDescriptor) << L", dwUnused2 "
                   << vbl::Hex((unsigned long)dwUnused2) << L", ppDst " << vbl::Hex((unsigned long)ppDst) << L", pNewObj "
                   << vbl::Hex((unsigned long)pNewObj);

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
EXPORT HRESULT __vbaRaiseEvent(DWORD unk1, DWORD unk2, int argCount, ...)
{
    va_list args;

    va_start(args, argCount);

    LOG(LOG_DEBUG) << L"pMe " << vbl::Hex((unsigned long)unk1) << L", dispId " << vbl::Hex((unsigned long)unk2)
                   << L", argCount = " << argCount;

    vba_VBVTable* pMe = (vba_VBVTable*)unk1;
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
EXPORT IUnknown* __stdcall __vbaObjSet(IUnknown** ppiuDest, IUnknown* piuSrc)
{

    LOG(LOG_DEBUG) << L"ppiuDest " << vbl::Hex((unsigned long)ppiuDest) << L", piuSrc " << vbl::Hex((unsigned long)piuSrc);

    if (*ppiuDest)
    {
        (*ppiuDest)->Release();
    }

    *ppiuDest = piuSrc;

    LOG(LOG_DEBUG) << L"returning " << vbl::Hex((unsigned long)*ppiuDest);

    return *ppiuDest;
} /* __vbaObjSet */

/**
 * @brief			Releases the ref of the destination IUnknown pointer (if not null), and sets it
 *					with the source pointer. Also adds a reference of the source IUnknown object.
 * @param			ppiuDest		Pointer to a IUnknown *. This is where the piuSrc value will be written.
 * @param			piuSrc			Pointer to a IUnknown.
 * @returns			*ppiuDest always.
 */
EXPORT IUnknown* __stdcall __vbaObjSetAddref(IUnknown** ppiuDest, IUnknown* piuSrc)
{

    LOG(LOG_DEBUG) << L"ppiuDest " << vbl::Hex((unsigned long)ppiuDest) << L", piuSrc " << vbl::Hex((unsigned long)piuSrc);

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
EXPORT void* __stdcall __vbaCastObj(IUnknown* piunkIn, REFIID riid)
{
    HRESULT hr;
    void* ppvObject = nullptr;

    LOG(LOG_DEBUG) << L"piunkIn " << vbl::Hex((unsigned long)piunkIn) << L", riid " << vbl::Guid(riid);

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
EXPORT HRESULT __stdcall rtcCreateObject2(VARIANTARG* pvargObject, BSTR bstrClassName, BSTR bstrServerName)
{

    LOG(LOG_DEBUG) << L"pvargObject " << vbl::Hex((unsigned long)pvargObject) << L", bstrClassName "
                   << vbl::Hex((unsigned long)bstrClassName) << L" ('" << vbl::Bstr(bstrClassName) << L"')"
                   << L", bstrServerName " << vbl::Hex((unsigned long)bstrServerName);

    HRESULT hr;
    CLSID rCLSID;

    hr = CLSIDFromProgIDEx(bstrClassName, &rCLSID);

    LOG(LOG_DEBUG) << L"CLSIDFromProgIDEx: " << vbl::Guid(rCLSID);

    if (!SUCCEEDED(hr))
    {
        vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
        return hr;
    }

    IUnknown* ppv;
    hr = CoCreateInstance(rCLSID, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&ppv));
    LOG(LOG_DEBUG) << L"CoCreateInstance = " << vbl::Hres(hr) << L", ppv = " << vbl::Hex((unsigned long)ppv);

    if (!SUCCEEDED(hr))
    {
        vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
        return hr;
    }

    IDispatch* disp;
    hr = ppv->QueryInterface(IID_PPV_ARGS(&disp));
    LOG(LOG_DEBUG) << L"ppv->QueryInterface = " << vbl::Hres(hr) << L", disp = " << vbl::Hex((unsigned long)disp);

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
EXPORT IDispatch* __stdcall __vbaObjVar(VARIANTARG* pvargIn)
{
    VARTYPE vtype;
    IDispatch* ret = nullptr;

    LOG(LOG_DEBUG) << L"pvargIn " << vbl::Hex((unsigned long)pvargIn);

    if (!pvargIn)
    {
        vbaRaiseException(VBA_EXCEPTION_OBJECT_REQUIRED);
        return nullptr;
    }

    LOG(LOG_DEBUG) << L"pvargIn type " << vbl::Hex((unsigned long)pvargIn->vt) << L", pvargIn plVal "
                   << vbl::Hex((unsigned long)pvargIn->plVal);

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
EXPORT void __cdecl __vbaLateMemCall(IDispatch* pidObject, BSTR bstrMethodName, int argCount, ...)
{
    HRESULT hr;

    LOG(LOG_DEBUG) << L"pidObject " << vbl::Hex((unsigned long)pidObject) << L", bstrMethodName "
                   << vbl::Hex((unsigned long)bstrMethodName) << L", argCount " << vbl::Hex((unsigned long)argCount);

    if (!pidObject)
    {
        vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
        return;
    }

    DISPID dispID;

    /* Get the ID of the method */
    hr = pidObject->GetIDsOfNames(IID_NULL, &bstrMethodName, 1, LOCALE_USER_DEFAULT, &dispID);

    LOG(LOG_DEBUG) << L"GetIDsOfNames = " << vbl::Hres(hr) << L", dispID = " << vbl::Hex((unsigned long)dispID);

    if (hr != S_OK)
    {
        LOG(LOG_DEBUG) << L"GetIDsOfNames failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError());

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right

        return;
    }

    va_list args;
    va_start(args, argCount);

    DISPPARAMS dispParamsInput;

    dispParamsInput.cArgs = argCount;
    dispParamsInput.rgvarg = (VARIANTARG*)args;

    va_end(args);

    dispParamsInput.cNamedArgs = 0;
    dispParamsInput.rgdispidNamedArgs = nullptr;

    EXCEPINFO excepInfo{};
    UINT uArgErr = 0;

    /* Invoke the method */
    hr = pidObject->Invoke(
        dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParamsInput, nullptr, &excepInfo, &uArgErr);

    LOG(LOG_DEBUG) << L"pidObject->Invoke = " << vbl::Hres(hr);

    if (!SUCCEEDED(hr))
    {
        LOG(LOG_DEBUG) << L"pidObject->Invoke failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError()) << L", excepInfo = '"
                       << vbl::Bstr(excepInfo.bstrDescription) << L"'" << L", uArgErr = " << vbl::Hex((unsigned long)uArgErr);

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR, &excepInfo); // TODO: Check if this exception is right
    }
} /* __vbaLateMemCall */

EXPORT void __stdcall __vbaVarLateMemSt(VARIANTARG* pvargObject, BSTR bstrMethodName, int argCount, ...)
{
    HRESULT hr;

    LOG(LOG_DEBUG) << L"Object " << vbl::Hex((unsigned long)pvargObject) << L", bstrMethodName "
                   << vbl::Hex((unsigned long)bstrMethodName) << L", argCount " << vbl::Hex((unsigned long)argCount);

    IDispatch* pidObject;

    /* Get the IDispatch object from the Variant */
    pidObject = __vbaObjVar(pvargObject);

    if (!pidObject)
    {
        vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
        return;
    }

    DISPID dispID;

    /* Get the ID of the method */
    hr = pidObject->GetIDsOfNames(IID_NULL, &bstrMethodName, 1, LOCALE_USER_DEFAULT, &dispID);

    LOG(LOG_DEBUG) << L"GetIDsOfNames = " << vbl::Hres(hr) << L", dispID = " << vbl::Hex((unsigned long)dispID);

    if (hr != S_OK)
    {
        LOG(LOG_DEBUG) << L"GetIDsOfNames failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError());

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right

        return;
    }

    va_list args;
    va_start(args, argCount);

    DISPPARAMS dispParamsInput;

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
EXPORT VARIANTARG* __cdecl __vbaVarLateMemCallLdRf(VARIANTARG* pvargRet, VARIANTARG* pvargObject, BSTR bstrMethodName, int argCount, ...)
{
    HRESULT hr;

    LOG(LOG_DEBUG) << L"pvargRet " << vbl::Hex((unsigned long)pvargRet) << L", pvargRet->vt " << vbl::Hex((unsigned long)pvargRet->vt)
                   << L", pvargObject " << vbl::Hex((unsigned long)pvargObject) << L", bstrMethodName "
                   << vbl::Hex((unsigned long)bstrMethodName) << L", argCount " << vbl::Hex((unsigned long)argCount);

    if (!pvargObject)
    {
        vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
        return pvargRet;
    }

    IDispatch* pidObject;

    /* Get the IDispatch object from the Variant */
    pidObject = __vbaObjVar(pvargObject);

    if (!pidObject)
    {
        LOG(LOG_DEBUG) << L"Could not get IDispatch object from the pvargObject " << vbl::Hex((unsigned long)pvargObject);

        return pvargRet;
    }

    DISPID dispID;

    /* Get the ID number of the requested method */
    hr = pidObject->GetIDsOfNames(IID_NULL, &bstrMethodName, 1, LOCALE_USER_DEFAULT, &dispID);

    LOG(LOG_DEBUG) << L"GetIDsOfNames = " << vbl::Hres(hr) << L", dispID = " << vbl::Hex((unsigned long)dispID);

    if (hr != S_OK)
    {
        LOG(LOG_DEBUG) << L"GetIDsOfNames failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError());

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right

        return pvargRet;
    }

    va_list args;
    va_start(args, argCount);

    DISPPARAMS dispParamsInput;

    dispParamsInput.cArgs = argCount;
    dispParamsInput.rgvarg = (VARIANTARG*)args;

    LOG(LOG_DEBUG) << L"dispParamsInput.rgvarg " << vbl::Hex((unsigned long)dispParamsInput.rgvarg);

    va_end(args);

    dispParamsInput.cNamedArgs = 0;
    dispParamsInput.rgdispidNamedArgs = nullptr;

    EXCEPINFO excepInfo;
    UINT uArgErr;

    /* Invoke the method */
    hr = pidObject->Invoke(
        dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParamsInput, pvargRet, &excepInfo, &uArgErr);

    LOG(LOG_DEBUG) << L"pidObject->Invoke = " << vbl::Hres(hr) << L", pvargRet->vt " << vbl::Hex((unsigned long)pvargRet->vt);

    if (hr != S_OK)
    {
        LOG(LOG_DEBUG) << L"pidObject->Invoke failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError()) << L", excepInfo = '"
                       << vbl::Bstr(excepInfo.bstrDescription) << L"'" << L", uArgErr = " << vbl::Hex((unsigned long)uArgErr);

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right
    }

    return pvargRet;
} /* __vbaVarLateMemCallLdRf */

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
HRESULT objIDispatchGetDefaultValue(IDispatch* pidObject, VARIANTARG* pvargValueOut)
{
    HRESULT hr;

    LOG(LOG_DEBUG) << L"pidObject " << vbl::Hex((unsigned long)pidObject) << L", pvargValueOut "
                   << vbl::Hex((unsigned long)pvargValueOut);

    if (!pidObject)
    {
        vbaRaiseException(VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET);
        return -1;
    }

    DISPPARAMS dispParamsInput = {0};
    EXCEPINFO excepInfo;
    UINT uArgErr;

    /* Invoke the method */
    hr = pidObject->Invoke(
        0, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dispParamsInput, pvargValueOut, &excepInfo, &uArgErr);

    LOG(LOG_DEBUG) << L"pidObject->Invoke = " << vbl::Hres(hr) << L", pvargValueOut->vt " << vbl::Hex((unsigned long)pvargValueOut->vt);

    if (hr != S_OK)
    {
        LOG(LOG_DEBUG) << L"pidObject->Invoke failed! GetLastError() = " << vbl::Hex((unsigned long)GetLastError()) << L", excepInfo = '"
                       << vbl::Bstr(excepInfo.bstrDescription) << L"'" << L", uArgErr = " << vbl::Hex((unsigned long)uArgErr);

        vbaRaiseException(VBA_EXCEPTION_AUTOMATION_ERROR); // TODO: Check if this exception is right
    }

    return hr;
} /* objIDispatchGetDefaultValue */

/**
 * @brief			TBD
 * @param			TBD				TBD
 * @returns			TBD
 */
EXPORT VARIANTARG* __cdecl __vbaVarLateMemCallLd(VARIANTARG* pvargRet, VARIANTARG* pvarObject, BSTR bstrMethodName, int argCount, ...)
{

    LOG(LOG_DEBUG) << L"bridgeing to __vbaVarLateMemCallLdRf but it's wrong and will crash!";

    va_list args;
    va_start(args, argCount);
    return __vbaVarLateMemCallLdRf(pvargRet, pvarObject, bstrMethodName, argCount, args);
} /* __vbaVarLateMemCallLd */

/**
 * @brief			Releases a COM Object (IUnknown) via its pointer, and nulls it.
 * @param			ppunkObj		Pointer to an object to release, and null the pointer.
 * @returns			Dereferenced object of ppunkObj.
 */
EXPORT void __fastcall __vbaFreeObj(IUnknown** punkObj)
{

    LOG(LOG_TRACE) << L"punkObj " << vbl::Hex(punkObj);

    if (punkObj)
    {
        LOG(LOG_TRACE) << L"*punkObj " << vbl::Hex(*punkObj);

        if (*punkObj)
        {
            (*punkObj)->Release();
            *punkObj = nullptr;
        }
    }
} /* __vbaFreeObj */

/**
 * @brief			Releases a list of COM Objects (IUnknowns) via their pointers, and nulls them.
 * @param			argCount		Count of elements.
 * @param			...				Pointers to objects to release and null them.
 */
EXPORT void __cdecl __vbaFreeObjList(unsigned int uiArgCount, ...)
{
    IUnknown** piunkElement;

    va_list args;
    va_start(args, uiArgCount);

    while (uiArgCount--)
    {
        piunkElement = va_arg(args, IUnknown**);
        LOG(LOG_DEBUG) << L"arg() " << vbl::Hex((unsigned long)piunkElement) << L", argsRemaining "
                       << vbl::Hex((unsigned long)uiArgCount);

        __vbaFreeObj(piunkElement);
    }

    va_end(args);
} /* __vbaFreeObjList */
