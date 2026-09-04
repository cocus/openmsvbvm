#pragma once

#include "vba_internal.h"
#include <ocidl.h>
#include <olectl.h>
#include <vector>

#include "vba_structures.h"

/**
 * One advised sink on a vbConnectionPoint: the (already QueryInterface'd/AddRef'd)
 * IDispatch to call, and the cookie handed back to the caller of Advise().
 */
struct SinkConnection
{
	IDispatch	*pSink;
	DWORD		dwCookie;
};

/**
 * Minimal, real IConnectionPoint implementation backing a class's outgoing event
 * interface. Lifetime is tied to its owning vbClassWrapper (it's held as a value
 * member, not separately heap-allocated), so AddRef/Release are no-ops rather than
 * an independent refcount. Always answers as the (single) connection point for
 * whatever interface FindConnectionPoint on the container was asked for -- a class
 * with more than one distinct outgoing interface isn't disambiguated (see plan's
 * Known Limitations).
 */
class vbConnectionPoint : public IConnectionPoint
{
public:
	explicit vbConnectionPoint(IUnknown *pContainer);

	// IUnknown
	HRESULT __stdcall QueryInterface(REFIID riid, void **ppv) override;
	ULONG   __stdcall AddRef() override;
	ULONG   __stdcall Release() override;

	// IConnectionPoint
	HRESULT __stdcall GetConnectionInterface(IID *pIID) override;
	HRESULT __stdcall GetConnectionPointContainer(IConnectionPointContainer **ppCPC) override;
	HRESULT __stdcall Advise(IUnknown *pUnkSink, DWORD *pdwCookie) override;
	HRESULT __stdcall Unadvise(DWORD dwCookie) override;
	HRESULT __stdcall EnumConnections(IEnumConnections **ppEnum) override;

	std::vector<SinkConnection> & Sinks() { return m_sinks; }

private:
	IUnknown						*m_pContainer; // weak back-pointer, not addref'd
	std::vector<SinkConnection>	m_sinks;
	DWORD							m_nextCookie = 1;
};

/**
 * A single connected event sink: wraps the compiler-built, per-class "handler thunk
 * array" block (shared/static -- see FindEventSinkBlock) together with the specific
 * live owner instance ("Me") this particular connection was made for, captured at
 * Advise time. This is what actually gets handed to IConnectionPoint::Advise -- the
 * shared static block itself is never advised directly, since it can't tell two
 * different owning instances apart.
 */
class vbEventSinkInstance : public IDispatch
{
public:
	vbEventSinkInstance(IDispatch *pStaticSinkBlock, void *pOwnerMe);

	HRESULT __stdcall QueryInterface(REFIID riid, void **ppv) override;
	ULONG   __stdcall AddRef() override;
	ULONG   __stdcall Release() override;

	HRESULT __stdcall GetTypeInfoCount(UINT *pctInfo) override;
	HRESULT __stdcall GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo **pptinfo) override;
	HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgdispid) override;
	HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr) override;

private:
	long		m_nRefCount;
	IDispatch	*m_pStaticSinkBlock;
	void		*m_pOwnerMe;
};

/**
 * Walks a connection point's sinks and calls Invoke(dispId, ...) on each, matching
 * the real RaiseEventOnBasicClass's Invoke call shape exactly (lcid 0x409, wFlags=1,
 * DISPPARAMS built from pArgs/argCount, no named args, no return value).
 */
HRESULT RaiseEventOnSinks(
	std::vector<SinkConnection>	&sinks,
	DISPID						dispId,
	VARIANTARG					*pArgs,
	DWORD						argCount
);

/**
 * Signature-scans a compiled class's static "Controls" data (opt.lpControls through
 * opt.lpEvents) for a compiler-built WithEvents sink block: a real 7-slot IDispatch
 * vtable (EVENT_SINK_QueryInterface/AddRef/Release, then real EVENT_SINK_GetIDsOfNames/
 * Invoke -- not thunks, which is what distinguishes a genuine WithEvents sink from the
 * similarly-shaped "self late-bound surface" blocks the compiler also emits) followed
 * by an embedded array of per-dispId handler thunks. If pTargetDescriptor is given,
 * prefers a sink whose referencing control entry also references that class's own
 * descriptor (disambiguates multiple WithEvents variables); otherwise/on no match,
 * falls back to the first sink found (correct for the common single-WithEvents-
 * variable case).
 */
IDispatch * FindEventSinkBlock(
	ObjectInfoWithOptional	*pOwnerDescriptor,
	ObjectInfoWithOptional	*pTargetDescriptor
);

/** Reads the dispId-th (1-based) handler thunk embedded after a sink block's vtable. */
void * GetHandlerThunk(IDispatch *pSinkBlock, DISPID dispId);

/**
 * Parses a handler thunk's leading "sub dword ptr [esp+4], imm" instruction (either
 * the imm8 or imm32 encoding -- both observed across different classes/methods in
 * this session) and returns the immediate value.
 */
int ReadThunkAdjustment(void *pThunk);

/**
 * Resolves a handler thunk's own "jmp Handler" target address (skipping past its
 * leading "sub esp+4,imm" instruction first) -- i.e. the real compiled Sub's entry
 * point, e.g. "Form1::Form_QueryUnload". Returns nullptr if pThunk doesn't look like
 * the recognized thunk shape.
 */
void * ResolveHandlerThunkTarget(void *pThunk);

/**
 * Best-effort stdcall argument-byte-count classifier for a resolved handler target
 * (see ResolveHandlerThunkTarget) -- used to tell which intrinsic event a
 * declaration-order-packed thunk is when a class implements more than one (no
 * compiled name/ordinal table exists for this -- see EventDispatch.cpp). Returns -1
 * if no recognizable epilogue is found.
 */
int GetStdcallArgBytes(void *pFuncStart);

/**
 * Computes (pOwnerMe + ReadThunkAdjustment(pThunk)) and calls pThunk with that as the
 * implicit first (stdcall) argument, unpacking pArgs as the remaining ByVal
 * arguments. Supports the common ByVal primitive VARIANT shapes (4-byte types plus
 * 8-byte Double/Currency/Date); ByRef/array/UDT/object marshaling is out of scope.
 */
void InvokeHandlerThunk(
	void		*pThunk,
	void		*pOwnerMe,
	VARIANTARG	*pArgs,
	DWORD		argCount
);

/**
 * Tracks the "Me" of whichever vbClassWrapper-driven method (Class_Initialize/
 * Class_Terminate) is currently executing on this thread, so a WithEvents Advise
 * happening inside one of those can capture the owning instance for later handler
 * invocation. Known limitation: a WithEvents assignment made from an arbitrary other
 * method (reached directly by compiled code via a vtable offset, bypassing this
 * project's own C++ code entirely) won't have a current-instance context pushed.
 */
void PushCurrentInstance(void *pMe);
void PopCurrentInstance();
void * GetCurrentInstance();

struct CurrentInstanceScope
{
	explicit CurrentInstanceScope(void *pMe) { PushCurrentInstance(pMe); }
	~CurrentInstanceScope() { PopCurrentInstance(); }
};
