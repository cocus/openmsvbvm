#pragma once

#include <windows.h>

/**
 * Everything CreateFormWindow needs from a compiled Form's parsed template
 * (FormProperties.hpp's FormTemplateProps) plus its owner, kept as plain fields here
 * so this header doesn't need to depend on FormProperties.hpp.
 */
struct FormWindowCreateParams
{
	const wchar_t	*caption;
	void			*pOwnerVBVTable;

	// twips; LONG_MIN on either position field means "let Windows place it"
	// (FormProperties.hpp's InitFormTemplateDefaults sentinel) -- width/height are
	// always concrete.
	LONG			clientLeftTwips;
	LONG			clientTopTwips;
	LONG			clientWidthTwips;
	LONG			clientHeightTwips;

	DWORD			backColor; // raw OLE_COLOR (0x00BBGGRR, or 0x800000xx for a system color index)
	BYTE			borderStyle; // 0=None,1=Fixed Single,2=Sizable,3=Fixed Dialog,4=Fixed ToolWindow,5=Sizable ToolWindow

	bool			enabled;
	bool			controlBox;
	bool			maxButton;
	bool			minButton;
};

/**
 * Creates a real top-level window for a VB6 Form instance, registering the shared
 * "ThunderRT6FormDC" window class (the real msvbvm60 name, confirmed via disassembly)
 * on first use. pHwndSlot is the caller's own HWND storage (e.g. a vbFormWrapper's
 * m_hwnd field): it's written with the new HWND immediately, and stashed via
 * GWLP_USERDATA so the shared window procedure can null it back out on WM_DESTROY
 * (whether the window was destroyed by us or by the user closing it). params.
 * pOwnerVBVTable is the owning object's vba_VBVTable* (opaque here), stashed as a
 * window property so WM_CLOSE can attempt Form_QueryUnload dispatch (see
 * ObjectManipulation.cpp's VBFormTryQueryUnload) before actually closing. Returns the
 * new HWND, or nullptr on failure.
 */
HWND CreateFormWindow(
	HWND						*pHwndSlot,
	const FormWindowCreateParams	&params
);

/**
 * Runs a local modal message loop (GetMessage/TranslateMessage/DispatchMessage)
 * until *pHwndSlot becomes null -- i.e. until the window tracked by that slot is
 * destroyed, whether via WM_CLOSE's default DestroyWindow or any other path that
 * ends in WM_DESTROY.
 */
void RunModalMessageLoop(
	HWND			*pHwndSlot
);
