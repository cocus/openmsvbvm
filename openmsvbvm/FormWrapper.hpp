#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "ClassWrapper.hpp"

/**
 * Form-specific half of the Form/plain-class split (see ClassWrapper.hpp's header
 * comment): gives a Form-derived VB6 class instance a real Win32 window and dispatches
 * its intrinsic events (Load/Show/QueryUnload). Detached from vbClassWrapper so that
 * class carries no Form concept at all -- __vbaNew (ObjectManipulation.cpp)
 * constructs one of these instead of a plain vbClassWrapper whenever
 * IsFormLikeDescriptor() says the class being instantiated is a Form.
 */
class vbFormWrapper : public vbClassWrapper
{
public:
	vbFormWrapper(vba_VBVTable * pWrapperVtable, ObjectInfoWithOptional* pObjInfo);
	~vbFormWrapper() override;

	// Real Load/Show/Unload backing for Form-derived objects (see FormWindow.hpp).
	bool EnsureWindowCreated();
	void DestroyFormWindow();
	HRESULT Show(VARIANTARG windowStyle, VARIANTARG ownerForm);
	bool TryFireQueryUnload(short *pCancel);
	void TryFireLoadOnce();

	// The window itself -- _FormImpl (FormWrapper.cpp) is the one place that actually
	// implements _Form's COM property surface (get_Caption, etc.), calling straight
	// through to Win32 APIs against this handle; this class only owns the handle's
	// lifecycle (creation/destruction/lookup), not any COM-shaped reflection of it.
	HWND GetHwnd() const { return m_hwnd; }

private:
	HWND						m_hwnd = nullptr;
	BYTE						m_windowState = 0; // 0=Normal,1=Minimized,2=Maximized (FormProperties.hpp's WindowState)
	bool						m_bLoadFired = false;

	void *FindIntrinsicThunkByArgBytes(int wantedArgBytes) const;
};

/**
 * Returns the real, compiler-generated `_Form` vtable (Form.idl, MIDL-compiled --
 * see FormWrapper.cpp's _FormImpl, a genuine C++ class implementing `_Form` the same
 * way App.cpp's _AppImpl implements `_App`), starting right after its inherited
 * IDispatch slots (QueryInterface/AddRef/Release/GetTypeInfoCount/GetTypeInfo/
 * GetIDsOfNames/Invoke -- this project's own vba_BASIC_CLASS_IUnknownBridge already
 * covers those 7 slots for every wrapped object, Form or not), and writes the number
 * of DWORD-sized slots spanned from there through `_Form`'s last member (OLEDrag) to
 * *pSlotCount.
 *
 * __vbaNew (ObjectManipulation.cpp) memcpy's this straight into a Form-derived
 * class's synthesized per-instance vtable immediately after its own 7-slot bridge --
 * landing every real `_Form` member at its true, compiler-verified offset with no
 * manual per-member offset math anywhere in this project (Show's real, disassembly-
 * confirmed 0x2B0 byte offset falls out of this automatically: it's exactly
 * sizeof(vba_BASIC_CLASS_IUnknownBridge) plus Show's own offset from the start of
 * this returned tail, and both of those are now compiler facts, not constants this
 * project maintains by hand).
 */
void * const * GetFormInterfaceVtableTail(size_t *pSlotCount);
