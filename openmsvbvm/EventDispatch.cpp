#include "Logging.hpp"
#include "EventDispatch.hpp"

#include <cstring>

/* These are the same generic bridge exports ObjectManipulation.cpp defines for a
   compiler-built WithEvents sink block's vtable slots -- declared here again (they
   match those definitions exactly) so this file can take their addresses for the
   signature scan in FindEventSinkBlock, without needing to expose them from a
   shared header. */
extern "C" HRESULT __stdcall EVENT_SINK_QueryInterface(void* arg1, REFIID arg2, void** ppvObject);
extern "C" HRESULT __stdcall EVENT_SINK_AddRef(void* arg1);
extern "C" HRESULT __stdcall EVENT_SINK_Release(void* arg1);
extern "C" HRESULT __stdcall EVENT_SINK_GetIDsOfNames(OLECHAR* strIn, DWORD unk1, DWORD unk2, DWORD unk3, DWORD unk4, DWORD unk5);
extern "C" HRESULT __stdcall EVENT_SINK_Invoke(
    DWORD* unk1, DWORD unk2, DWORD unk3, DWORD unk4, DWORD unk5, DWORD unk6, DWORD unk7, DWORD unk8, DWORD unk9);

//
// vbConnectionPoint
//

vbConnectionPoint::vbConnectionPoint(IUnknown* pContainer) : m_pContainer(pContainer)
{
}

HRESULT __stdcall vbConnectionPoint::QueryInterface(REFIID riid, void** ppv)
{
    if (!ppv)
    {
        return E_POINTER;
    }

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IConnectionPoint))
    {
        *ppv = static_cast<IConnectionPoint*>(this);
        AddRef();
        return S_OK;
    }

    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall vbConnectionPoint::AddRef()
{
    return 1;
}

ULONG __stdcall vbConnectionPoint::Release()
{
    return 1;
}

HRESULT __stdcall vbConnectionPoint::GetConnectionInterface(IID* pIID)
{
    if (!pIID)
    {
        return E_POINTER;
    }

    // TODO: resolve the real outgoing-interface IID via tagRegInfo::bUuidEventsIFace
    // (see plan's Known Limitations -- not wired up yet).
    *pIID = IID_NULL;
    return S_OK;
}

HRESULT __stdcall vbConnectionPoint::GetConnectionPointContainer(IConnectionPointContainer** ppCPC)
{
    if (!ppCPC)
    {
        return E_POINTER;
    }

    if (!m_pContainer)
    {
        *ppCPC = nullptr;
        return E_NOINTERFACE;
    }

    return m_pContainer->QueryInterface(IID_IConnectionPointContainer, (void**)ppCPC);
}

HRESULT __stdcall vbConnectionPoint::Advise(IUnknown* pUnkSink, DWORD* pdwCookie)
{

    if (!pUnkSink || !pdwCookie)
    {
        return E_POINTER;
    }

    IDispatch* pSinkDispatch = nullptr;
    HRESULT hr = pUnkSink->QueryInterface(IID_IDispatch, (void**)&pSinkDispatch);
    if (FAILED(hr) || !pSinkDispatch)
    {
        return CONNECT_E_CANNOTCONNECT;
    }

    SinkConnection conn;
    conn.pSink = pSinkDispatch; // holds the reference QueryInterface just returned
    conn.dwCookie = m_nextCookie++;
    m_sinks.push_back(conn);

    *pdwCookie = conn.dwCookie;

    LOG(LOG_DEBUG) << L"advised sink " << vbl::Hex((unsigned long)pSinkDispatch) << L", cookie "
                   << vbl::Hex((unsigned long)conn.dwCookie);

    return S_OK;
}

HRESULT __stdcall vbConnectionPoint::Unadvise(DWORD dwCookie)
{

    for (auto it = m_sinks.begin(); it != m_sinks.end(); ++it)
    {
        if (it->dwCookie == dwCookie)
        {
            LOG(LOG_DEBUG) << L"unadvising sink " << vbl::Hex((unsigned long)it->pSink) << L", cookie "
                           << vbl::Hex((unsigned long)dwCookie);

            if (it->pSink)
            {
                it->pSink->Release();
            }

            m_sinks.erase(it);
            return S_OK;
        }
    }

    return CONNECT_E_NOCONNECTION;
}

HRESULT __stdcall vbConnectionPoint::EnumConnections(IEnumConnections** ppEnum)
{
    return E_NOTIMPL;
}

//
// vbEventSinkInstance
//

vbEventSinkInstance::vbEventSinkInstance(IDispatch* pStaticSinkBlock, void* pOwnerMe) :
    m_nRefCount(1), m_pStaticSinkBlock(pStaticSinkBlock), m_pOwnerMe(pOwnerMe)
{
}

HRESULT __stdcall vbEventSinkInstance::QueryInterface(REFIID riid, void** ppv)
{
    if (!ppv)
    {
        return E_POINTER;
    }

    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
    {
        *ppv = static_cast<IDispatch*>(this);
        AddRef();
        return S_OK;
    }

    *ppv = nullptr;
    return E_NOINTERFACE;
}

ULONG __stdcall vbEventSinkInstance::AddRef()
{
    return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall vbEventSinkInstance::Release()
{
    long nRefCount = InterlockedDecrement(&m_nRefCount);
    if (nRefCount == 0)
    {
        delete this;
    }
    return nRefCount;
}

HRESULT __stdcall vbEventSinkInstance::GetTypeInfoCount(UINT* pctInfo)
{
    if (pctInfo)
    {
        *pctInfo = 0;
    }
    return S_OK;
}

HRESULT __stdcall vbEventSinkInstance::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
{
    return TYPE_E_ELEMENTNOTFOUND;
}

HRESULT __stdcall vbEventSinkInstance::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid)
{
    return E_NOTIMPL;
}

HRESULT __stdcall vbEventSinkInstance::Invoke(
    DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{

    void* pThunk = GetHandlerThunk(m_pStaticSinkBlock, dispIdMember);

    LOG(LOG_DEBUG) << L"dispIdMember " << dispIdMember << L", pOwnerMe " << vbl::Hex((unsigned long)m_pOwnerMe) << L", pThunk "
                   << vbl::Hex((unsigned long)pThunk);

    if (!pThunk)
    {
        return S_OK;
    }

    InvokeHandlerThunk(pThunk, m_pOwnerMe, pDispParams ? pDispParams->rgvarg : nullptr, pDispParams ? pDispParams->cArgs : 0);

    return S_OK;
}

//
// RaiseEventOnSinks
//

HRESULT RaiseEventOnSinks(std::vector<SinkConnection>& sinks, DISPID dispId, VARIANTARG* pArgs, DWORD argCount)
{

    // Snapshot + addref the live sinks before firing, so a handler that disconnects
    // itself (or another sink) mid-fire doesn't invalidate the list we're iterating.
    std::vector<IDispatch*> snapshot;
    snapshot.reserve(sinks.size());
    for (auto& conn : sinks)
    {
        if (conn.pSink)
        {
            conn.pSink->AddRef();
            snapshot.push_back(conn.pSink);
        }
    }

    DISPPARAMS dispParams;
    dispParams.rgvarg = argCount ? pArgs : nullptr;
    dispParams.rgdispidNamedArgs = nullptr;
    dispParams.cArgs = argCount;
    dispParams.cNamedArgs = 0;

    for (IDispatch* pSink : snapshot)
    {
        EXCEPINFO excepInfo{};
        UINT argErr = 0;

        // Matches the real RaiseEventOnBasicClass's Invoke call exactly: lcid 0x409,
        // wFlags=DISPATCH_METHOD, no return value.
        HRESULT hr = pSink->Invoke(dispId, IID_NULL, 0x409, DISPATCH_METHOD, &dispParams, nullptr, &excepInfo, &argErr);

        LOG(LOG_DEBUG) << L"pSink " << vbl::Hex((unsigned long)pSink) << L", dispId " << dispId << L", hr " << vbl::Hres(hr);

        pSink->Release();
    }

    return S_OK;
}

//
// FindEventSinkBlock
//

/**
 * A compiled EXE can't embed a literal address from a *different* module (this DLL)
 * as a compile-time constant -- the value it actually stores as "EVENT_SINK_QueryInterface"
 * etc. is the address of a small local import thunk (either a direct "jmp rel32", e.g.
 * an incremental-linking ILT-style stub, or an indirect "jmp dword ptr [iatSlot]"
 * through its own Import Address Table).
 */
static void* ResolveThunk(void* pAddr)
{
    if (!pAddr)
    {
        return nullptr;
    }

    BYTE* p = (BYTE*)pAddr;

    __try
    {
        if (p[0] == 0xE9) // jmp rel32
        {
            int rel;
            memcpy(&rel, p + 1, sizeof(rel));
            return p + 5 + rel;
        }

        if (p[0] == 0xFF && p[1] == 0x25) // jmp dword ptr [addr]
        {
            void** pIatSlot;
            memcpy(&pIatSlot, p + 2, sizeof(pIatSlot));
            return pIatSlot ? *pIatSlot : nullptr;
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return pAddr;
    }

    return pAddr;
}

static bool PointsToSameFunction(void* pStoredValue, void* pOurFunction)
{
    // pOurFunction is compared as-is, NOT further resolved: under incremental linking
    // (Debug builds here), taking &SomeExportedFunction from inside this DLL gives an
    // ILT thunk address, and that is the "stable" address the linker points this DLL's
    // own export table at, which is in turn what the compiled EXE's Import Address Table
    // slot gets patched to at load time.
    // pStoredValue is what the compiled EXE embedded for a cross-module function
    // reference: its own local import thunk, needing exactly one level of resolving
    // (see ResolveThunk) to reach that same "as-exported" address.
    return ResolveThunk(pStoredValue) == pOurFunction;
}

static bool IsRealEventSinkBlock(BYTE* pQueryInterfaceSlot)
{
    // pQueryInterfaceSlot points at the block's QueryInterface vtable slot; the block
    // itself (2 header dwords + QueryInterface) starts 0xC bytes earlier. Genuine
    // WithEvents sinks have the real generic EVENT_SINK_GetIDsOfNames/EVENT_SINK_Invoke
    // exports at +0x20/+0x24 (relative to the block start); the similarly-shaped
    // "self late-bound surface" blocks the compiler also builds use Zombie_* stand-ins
    // or local per-method thunks there instead.
    BYTE* pBlock = pQueryInterfaceSlot - 0xC;
    void* slotGetIDsOfNames = *(void**)(pBlock + 0x20);
    void* slotInvoke = *(void**)(pBlock + 0x24);

    return PointsToSameFunction(slotGetIDsOfNames, (void*)EVENT_SINK_GetIDsOfNames) &&
           PointsToSameFunction(slotInvoke, (void*)EVENT_SINK_Invoke);
}

IDispatch* FindEventSinkBlock(ObjectInfoWithOptional* pOwnerDescriptor, ObjectInfoWithOptional* pTargetDescriptor)
{

    if (!pOwnerDescriptor || !pOwnerDescriptor->opt.lpControls)
    {
        LOG(LOG_DEBUG) << L"owner descriptor has no Controls table";
        return nullptr;
    }

    BYTE* pStart = (BYTE*)pOwnerDescriptor->opt.lpControls;

    // The compiler-built sink blocks were found to sit AFTER opt.lpEvents in the
    // compiled .data layout (lpControls -> "control" entries -> lpEvents -> sink
    // block(s)), not before it, so lpEvents isn't a usable upper bound. There's no
    // independently-known end to this region, so scan a generous fixed window
    // forward from lpControls instead (the strict 3-function-address signature match
    // below makes false positives from unrelated data essentially impossible, so a
    // wide window is safe. The only real risk is scanning past the module's own
    // .data section, which 8 KB forward from a valid metadata pointer is very
    // unlikely to reach).
    BYTE* pEnd = pStart + 0x2000;

    IDispatch* pFirstMatch = nullptr;

    for (BYTE* p = pStart; p + 0x28 <= pEnd; p += sizeof(void*))
    {
        void** pDwords = (void**)p;

        if (!PointsToSameFunction(pDwords[0], (void*)EVENT_SINK_QueryInterface) ||
            !PointsToSameFunction(pDwords[1], (void*)EVENT_SINK_AddRef) || !PointsToSameFunction(pDwords[2], (void*)EVENT_SINK_Release))
        {
            continue;
        }

        if (!IsRealEventSinkBlock(p))
        {
            continue; // the class's own self-referencing late-bound surface, not a WithEvents sink
        }

        IDispatch* pCandidate = (IDispatch*)(p - 0xC);

        LOG(LOG_DEBUG) << L"candidate WithEvents sink block " << vbl::Hex((unsigned long)pCandidate);

        if (!pFirstMatch)
        {
            pFirstMatch = pCandidate;
        }

        if (!pTargetDescriptor)
        {
            continue;
        }

        // Cross-check: the "control" entry that references this sink also references
        // the target (source) class's own descriptor 4 bytes after wherever the sink
        // pointer itself is stored (confirmed for the reference test case). Scan the
        // same region for a pointer equal to this sink block's own address to find
        // that referencing entry.
        for (BYTE* q = pStart; q + 8 <= pEnd; q += sizeof(void*))
        {
            if (*(void**)q == (void*)pCandidate && *(void**)(q + 4) == (void*)pTargetDescriptor)
            {
                LOG(LOG_DEBUG) << L"matched sink " << vbl::Hex((unsigned long)pCandidate) << L" to target descriptor "
                               << vbl::Hex((unsigned long)pTargetDescriptor);
                return pCandidate;
            }
        }
    }

    LOG(LOG_DEBUG) << L"no target-matched sink found, falling back to first candidate " << vbl::Hex((unsigned long)pFirstMatch);

    return pFirstMatch;
}

void* GetHandlerThunk(IDispatch* pSinkBlock, DISPID dispId)
{
    if (!pSinkBlock || dispId < 1)
    {
        return nullptr;
    }

    BYTE* pBase = (BYTE*)pSinkBlock;
    void** pThunkArray = (void**)(pBase + 0x28);
    return pThunkArray[dispId - 1];
}

//
// Fixed-slot event dispatch (ControlInfo.lpEventTable) -- see vba_structures.h's
// ControlInfo doc comment for the confirmed table shape.
//

ControlInfo* FindOwnFormControlInfo(ObjectInfoWithOptional* pOwnerDescriptor)
{
    if (!pOwnerDescriptor || !pOwnerDescriptor->opt.lpControls)
    {
        return nullptr;
    }

    __try
    {
        ControlInfo* pControls = pOwnerDescriptor->opt.lpControls;
        for (DWORD i = 0; i < pOwnerDescriptor->opt.dwControlCount; i++)
        {
            if (pControls[i].dwIndex == 0xFFFFFFFF)
            {
                return &pControls[i];
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return nullptr;
    }

    return nullptr;
}

ControlInfo* FindPlacedControlByAccessorIndex(ObjectInfoWithOptional* pOwnerDescriptor, int accessorIndex)
{
    if (!pOwnerDescriptor || !pOwnerDescriptor->opt.lpControls || accessorIndex < 0)
    {
        return nullptr;
    }

    __try
    {
        /* opt.lpControls's own array STORAGE order is NOT what determines a
           control's accessor slot -- confirmed live it can change between
           otherwise-identical recompiles depending on which controls happen to be
           referenced by name in code (a build with only Label1 referenced stored
           it BEFORE the Form's own self entry; a later build with the same two
           controls stored it there too, but Command1 -- never referenced by name --
           ended up AFTER the self entry). What's stable across both is dwIndex's
           low WORD, a 1-based declaration/TabIndex-order value (confirmed live:
           Command1 always 1, Label1 always 2, matching their TabIndex+1) -- match
           accessorIndex (0-based) against THAT instead of raw array position. */
        ControlInfo* pControls = pOwnerDescriptor->opt.lpControls;
        for (DWORD i = 0; i < pOwnerDescriptor->opt.dwControlCount; i++)
        {
            if (pControls[i].dwIndex == 0xFFFFFFFF)
            {
                continue; // the Form/container's own entry, not a placed control
            }

            int declarationIndex = (int)(WORD)(pControls[i].dwIndex & 0xFFFF) - 1;
            if (declarationIndex == accessorIndex)
            {
                return &pControls[i];
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return nullptr;
    }

    return nullptr;
}

DWORD CountPlacedControls(ObjectInfoWithOptional* pOwnerDescriptor)
{
    if (!pOwnerDescriptor || !pOwnerDescriptor->opt.lpControls)
    {
        return 0;
    }

    DWORD count = 0;

    __try
    {
        ControlInfo* pControls = pOwnerDescriptor->opt.lpControls;
        for (DWORD i = 0; i < pOwnerDescriptor->opt.dwControlCount; i++)
        {
            if (pControls[i].dwIndex != 0xFFFFFFFF)
            {
                count++;
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return 0;
    }

    return count;
}

ControlInfo* FindControlInfoByName(ObjectInfoWithOptional* pOwnerDescriptor, const char* pszControlName)
{
    if (!pOwnerDescriptor || !pOwnerDescriptor->opt.lpControls || !pszControlName)
    {
        return nullptr;
    }

    __try
    {
        ControlInfo* pControls = pOwnerDescriptor->opt.lpControls;
        for (DWORD i = 0; i < pOwnerDescriptor->opt.dwControlCount; i++)
        {
            if (pControls[i].lpszName && strcmp(pControls[i].lpszName, pszControlName) == 0)
            {
                return &pControls[i];
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return nullptr;
    }

    return nullptr;
}

/* Same verification IsRealEventSinkBlock already does for a WithEvents sink block --
   a genuine lpEventTable has the real EVENT_SINK_QueryInterface/AddRef/Release
   exports at slots 3/4/5. */
static bool IsRealControlEventTable(void** pTable)
{
    return PointsToSameFunction(pTable[3], (void*)EVENT_SINK_QueryInterface) &&
           PointsToSameFunction(pTable[4], (void*)EVENT_SINK_AddRef) && PointsToSameFunction(pTable[5], (void*)EVENT_SINK_Release);
}

void* GetFixedEventThunk(LPVOID lpEventTable, int eventOrdinal)
{
    if (!lpEventTable || eventOrdinal < 0)
    {
        return nullptr;
    }

    const int kHeaderSlots = 6;

    __try
    {
        void** pTable = (void**)lpEventTable;

        if (!IsRealControlEventTable(pTable))
        {
            return nullptr;
        }

        return pTable[kHeaderSlots + eventOrdinal];
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return nullptr;
    }
}

int ReadThunkAdjustment(void* pThunkVoid)
{
    BYTE* pThunk = (BYTE*)pThunkVoid;
    if (!pThunk)
    {
        return 0;
    }

    // sub dword ptr [esp+4], imm8  (83 6C 24 04 ib)
    if (pThunk[0] == 0x83 && pThunk[1] == 0x6C && pThunk[2] == 0x24 && pThunk[3] == 0x04)
    {
        return (int)(signed char)pThunk[4];
    }

    // sub dword ptr [esp+4], imm32  (81 6C 24 04 id)
    if (pThunk[0] == 0x81 && pThunk[1] == 0x6C && pThunk[2] == 0x24 && pThunk[3] == 0x04)
    {
        int imm;
        memcpy(&imm, pThunk + 4, sizeof(imm));
        return imm;
    }

    // Unrecognized thunk shape -- best-effort: no adjustment.
    return 0;
}

typedef void(__stdcall* pFn0)(DWORD);
typedef void(__stdcall* pFn1)(DWORD, DWORD);
typedef void(__stdcall* pFn2)(DWORD, DWORD, DWORD);
typedef void(__stdcall* pFn3)(DWORD, DWORD, DWORD, DWORD);
typedef void(__stdcall* pFn4)(DWORD, DWORD, DWORD, DWORD, DWORD);
typedef void(__stdcall* pFn5)(DWORD, DWORD, DWORD, DWORD, DWORD, DWORD);
typedef void(__stdcall* pFn6)(DWORD, DWORD, DWORD, DWORD, DWORD, DWORD, DWORD);
typedef void(__stdcall* pFn7)(DWORD, DWORD, DWORD, DWORD, DWORD, DWORD, DWORD, DWORD);

void InvokeHandlerThunk(void* pThunk, void* pOwnerMe, VARIANTARG* pArgs, DWORD argCount)
{
    if (!pThunk)
    {
        LOG(LOG_WARN) << L"no thunk to invoke, pOwnerMe " << vbl::Hex(pOwnerMe);
        return;
    }

    int adjustment = ReadThunkAdjustment(pThunk);
    DWORD adjustedMe = (DWORD)pOwnerMe + (DWORD)adjustment;

    // Build the native argument buffer (VB-arg-order, arg1 first). Double/Currency/
    // Date are 8-byte on the x86 stdcall stack; the common ByVal primitive shapes
    // (String/Long/Integer/Boolean/Object/Single/...) are all 4-byte and safely
    // readable via the VARIANT union's .lVal (same first 4 bytes as bstrVal/
    // dispVal/punkVal/boolVal/etc.). ByRef/array/UDT marshaling is out of scope.
    DWORD argBuffer[8] = {0};
    int slotCount = 0;

    for (DWORD i = 0; i < argCount && slotCount < 7; i++)
    {
        VARTYPE vt = pArgs[i].vt & VT_TYPEMASK;

        if ((vt == VT_R8 || vt == VT_DATE || vt == VT_CY) && slotCount < 6)
        {
            unsigned __int64 wide;
            memcpy(&wide, &pArgs[i].dblVal, sizeof(wide));
            argBuffer[slotCount++] = (DWORD)(wide & 0xFFFFFFFFu);
            argBuffer[slotCount++] = (DWORD)(wide >> 32);
        }
        else
        {
            argBuffer[slotCount++] = (DWORD)pArgs[i].lVal;
        }
    }

    LOG(LOG_DEBUG) << L"pThunk " << vbl::Hex((unsigned long)pThunk) << L", pOwnerMe " << vbl::Hex((unsigned long)pOwnerMe)
                   << L", adjustment " << adjustment << L", adjustedMe " << vbl::Hex((unsigned long)adjustedMe) << L", argCount "
                   << argCount << L", slotCount " << slotCount;

    switch (slotCount)
    {
    case 0:
        ((pFn0)pThunk)(adjustedMe);
        break;
    case 1:
        ((pFn1)pThunk)(adjustedMe, argBuffer[0]);
        break;
    case 2:
        ((pFn2)pThunk)(adjustedMe, argBuffer[0], argBuffer[1]);
        break;
    case 3:
        ((pFn3)pThunk)(adjustedMe, argBuffer[0], argBuffer[1], argBuffer[2]);
        break;
    case 4:
        ((pFn4)pThunk)(adjustedMe, argBuffer[0], argBuffer[1], argBuffer[2], argBuffer[3]);
        break;
    case 5:
        ((pFn5)pThunk)(adjustedMe, argBuffer[0], argBuffer[1], argBuffer[2], argBuffer[3], argBuffer[4]);
        break;
    case 6:
        ((pFn6)pThunk)(adjustedMe, argBuffer[0], argBuffer[1], argBuffer[2], argBuffer[3], argBuffer[4], argBuffer[5]);
        break;
    case 7:
        ((pFn7)pThunk)(adjustedMe, argBuffer[0], argBuffer[1], argBuffer[2], argBuffer[3], argBuffer[4], argBuffer[5], argBuffer[6]);
        break;
    default:
        LOG(LOG_WARN) << L"too many argument slots (" << slotCount << L"), not calling handler";
        break;
    }
}

//
// Current-instance tracking
//

static thread_local std::vector<void*> g_currentInstanceStack;

void PushCurrentInstance(void* pMe)
{
    g_currentInstanceStack.push_back(pMe);
}

void PopCurrentInstance()
{
    if (!g_currentInstanceStack.empty())
    {
        g_currentInstanceStack.pop_back();
    }
}

void* GetCurrentInstance()
{
    if (g_currentInstanceStack.empty())
    {
        return nullptr;
    }
    return g_currentInstanceStack.back();
}
