#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "ControlWrapper.hpp"

#include <vector>

/**
 * Form-specific half of the Form/plain-class split (see ObjectWrapper.hpp's and
 * ControlWrapper.hpp's header comments): gives a Form-derived VB6 class instance a
 * real Win32 window and dispatches its intrinsic events (Load/Show/QueryUnload).
 * Derives from vbControlWrapper (this project's `CTL` equivalent), NOT from
 * vbClassWrapper (`CClassModule` equivalent) -- a Form is conceptually a control-like
 * visual object, not a plain class module with a window bolted on, matching how real
 * msvbvm60 keeps `FORM : public CTL` and `CClassModule` as unrelated hierarchies.
 * __vbaNew (ObjectManipulation.cpp) constructs one of these instead of a plain
 * vbClassWrapper whenever IsFormLikeDescriptor() says the class being instantiated is
 * a Form.
 */
class vbFormWrapper : public vbControlWrapper
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

	// Dispatches a WM_COMMAND notification (see FormWindow.cpp's FormWndProc) to
	// whichever placed control controlId identifies -- currently only a click
	// notification (BN_CLICKED for a button, STN_CLICKED for a label -- both
	// notification code 0), firing that control's Click handler via its own
	// fixed-slot event table entry (see ControlInfo's doc comment, vba_structures.h)
	// if implemented.
	void HandleCommand(WORD controlId, WORD notifyCode);

	// Returns a new "control accessor" object (see ControlObject.hpp) for the
	// accessorIndex-th REAL placed control in m_pObjInfo->opt.lpControls (0-based,
	// skipping the Form's own self entry -- see ControlInfo's doc comment,
	// vba_structures.h) -- what compiled code gets back from a direct vtable call
	// like "Label1.Caption = ..." (see __vbaNew, ObjectManipulation.cpp, for where
	// these accessor slots get reserved/wired up). Returns nullptr if accessorIndex
	// is out of range. The returned object is real regardless of whether this
	// project actually renders that control type -- its HWND is simply null if not,
	// so Caption reads back empty and silently no-ops on write in that case.
	// Qualified as ::IDispatch (not just IDispatch) -- this class's own base chain
	// privately inherits IDispatch (see ObjectWrapper.hpp), which would otherwise
	// make the unqualified name resolve to that private base subobject instead of
	// the unrelated, freshly-created object this method actually returns.
	::IDispatch * GetControlAccessor(int accessorIndex);

	// The window itself -- _FormImpl (FormWrapper.cpp) is the one place that actually
	// implements _Form's COM property surface (get_Caption, etc.), calling straight
	// through to Win32 APIs against this handle; this class only owns the handle's
	// lifecycle (creation/destruction/lookup), not any COM-shaped reflection of it.
	HWND GetHwnd() const { return m_hwnd; }

private:
	HWND						m_hwnd = nullptr;
	BYTE						m_windowState = 0; // 0=Normal,1=Minimized,2=Maximized (FormProperties.hpp's WindowState)
	bool						m_bLoadFired = false;

	// Set from ParseFormTemplate's own FormTemplateProps.pControlsAnchor
	// (EnsureWindowCreated) -- see that field's doc comment (FormProperties.hpp) for
	// why CreateChildControls must anchor each placed control's own
	// ParseControlTemplate scan here rather than on this class's raw
	// PublicObjectDescriptor.
	void						*m_pControlsAnchor = nullptr;

	// One entry per placed control this project knows how to actually create a real
	// window for (see ControlWindow.hpp -- CommandButton and Label so far). controlId
	// is this entry's own index into this vector, doubling as the child window's
	// Win32 control id (see ControlWindowCreateParams::controlId) so HandleCommand
	// can map a WM_COMMAND's wParam back to the right entry.
	struct ChildControl
	{
		HWND			hwnd;
		ControlInfo		*pControlInfo;    // for firing this control's own fixed-slot events
		int				clickEventOrdinal; // that control TYPE's own Click ordinal (see ControlWindow.hpp)
	};
	std::vector<ChildControl> m_controls;

	void CreateChildControls();
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
