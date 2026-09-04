#pragma once

#include <windows.h>

/**
 * Real msvbvm60 control-type codes, confirmed live via IDA against independently-
 * compiled calibration builds (see ControlInfo's own doc comment, vba_structures.h).
 * Only these two are confirmed so far.
 */
static const DWORD kControlTypeCommandButton = 0x00110040;
static const DWORD kControlTypeLabel = 0x00120040;

/**
 * Each control type's own real "XxxEvents" interface declares Click at a different
 * ordinal (see ControlInfo's doc comment for what "ordinal" means here) -- confirmed
 * live: CommandButtonEvents' Click is its 1st member (ordinal 0), LabelEvents' Click
 * is its 2nd (ordinal 1, something else -- unconfirmed what -- comes first).
 */
static const int kCommandButtonClickOrdinal = 0;
static const int kLabelClickOrdinal = 1;

struct ControlWindowCreateParams
{
	const wchar_t	*caption;

	// twips, matching the .frm source's own Left/Top/Width/Height keys
	LONG			leftTwips;
	LONG			topTwips;
	LONG			widthTwips;
	LONG			heightTwips;

	int				controlId; // becomes the child window's HMENU id, read back out of WM_COMMAND's wParam
};

/**
 * Creates a real Win32 BUTTON child window for a placed CommandButton, positioned/
 * sized against hwndParent's own client area (twips -> pixels at the real screen
 * DPI, same conversion FormWindow.cpp's CreateFormWindow already uses). Returns the
 * new HWND, or nullptr on failure.
 */
HWND CreateButtonControl(HWND hwndParent, const ControlWindowCreateParams &params);

/**
 * Creates a real Win32 STATIC child window for a placed Label, same positioning/
 * sizing as CreateButtonControl. SS_NOTIFY is set so it can report a click (a plain
 * STATIC control is otherwise silent) -- see kLabelClickOrdinal above for its Click
 * handler's dispatch ordinal (STN_CLICKED, like BN_CLICKED, is notification code 0,
 * so the existing WM_COMMAND handling needs no change to also cover this). Returns
 * the new HWND, or nullptr on failure.
 */
HWND CreateLabelControl(HWND hwndParent, const ControlWindowCreateParams &params);
