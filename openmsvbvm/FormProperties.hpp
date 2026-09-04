#pragma once

#include "vba_internal.h"
#include "vba_structures.h"

/**
 * Compiled-in designer properties for a Form-derived class, read straight out of the
 * compiled EXE's own data (twips, matching the .frm source) -- no .frx/resource file
 * involved. Every field is always valid: InitFormTemplateDefaults/ParseFormTemplate
 * fill in the real VB6 default for anything the compiled blob doesn't explicitly set
 * (VB6 only emits a tag for a property that differs from its default), so callers
 * never need a "has-this-field" fallback dance -- just use the field.
 *
 * Shape: one field per real `_Form` interface member (see C:\...\VB98\VB6.OLB,
 * confirmed via the same ITypeLib/ITypeInfo walk used to nail down Show's vtable
 * offset), in that interface's own declaration order, so this struct doubles as a
 * checklist of what a Form actually exposes. Left out entirely: `_Form`'s pure
 * runtime/live-state members that were never a compiled-in template value to begin
 * with -- hWnd, hDC, CurrentX/CurrentY, ActiveControl, Count, Controls -- since this
 * struct specifically models the *design-time* template, not live object state.
 *
 * Only a subset is actually parsed from the compiled blob today (see
 * ParseFormTemplate's tag table); OLE-object-typed members with no simple scalar
 * representation (Picture, Icon, MouseIcon, Font, Palette) are deliberately kept as
 * a bare LONG placeholder (0/-1 = "none set") rather than modeled as real objects --
 * real support would need genuine stdole.Picture/StdFont parsing, out of scope here.
 * Everything else not yet parsed still gets a field with its correct real VB6
 * default, ready for whenever its tag gets confirmed the same way the others were.
 */
struct FormTemplateProps
{
	char	name[64];      // Name -- the compiled class name (read-only at runtime in real VB6)
	char	caption[256];  // Caption -- PARSED (tag 0x0D)

	DWORD	backColor;     // BackColor -- PARSED (tag 0x03); raw OLE_COLOR (0x00BBGGRR or 0x800000xx)
	DWORD	foreColor;     // ForeColor -- default only; same OLE_COLOR encoding

	// Left/Top/Width/Height describe the whole form (including its frame) in the
	// real _Form interface; not yet parsed (only the separate ClientLeft/Top/Width/
	// Height quad below is confirmed) so these stay at a placeholder default.
	LONG	left;
	LONG	top;
	LONG	width;
	LONG	height;

	bool	enabled;       // Enabled -- PARSED (tag 0x09)
	BYTE	windowState;   // WindowState -- PARSED (tag 0x0A); 0=Normal,1=Minimized,2=Maximized

	BYTE	mousePointer;  // MousePointer -- default only

	// FontName/FontSize/... are _Form's own individual scalar properties (the real
	// underlying Font OLE object is the separate `font` placeholder further below,
	// matching how the interface exposes both).
	char	fontName[32];
	float	fontSize;
	bool	fontBold;
	bool	fontItalic;
	bool	fontStrikethru;
	bool	fontUnderline;
	bool	fontTransparent;

	// ScaleLeft/Top/Width/Height mirror Left/Top/Width/Height but in the form's
	// current ScaleMode units; not parsed, default only.
	LONG	scaleLeft;
	LONG	scaleTop;
	LONG	scaleWidth;
	LONG	scaleHeight;
	BYTE	scaleMode;

	BYTE	drawStyle;
	LONG	drawWidth;
	BYTE	fillStyle;
	DWORD	fillColor;
	BYTE	drawMode;
	bool	autoRedraw;

	LONG	picture;       // Picture -- IPictureDisp* in real VB6; opaque placeholder (0 = none)
	BYTE	borderStyle;   // BorderStyle -- PARSED (tag 0x22)
	LONG	icon;          // Icon -- PARSED presence only (tag 0x23); opaque placeholder (-1 = none set)

	char	linkTopic[256]; // LinkTopic -- PARSED (tag 0x24)
	BYTE	linkMode;

	bool	maxButton;     // PARSED (tag 0x26)
	bool	minButton;     // PARSED (tag 0x27)
	bool	controlBox;    // PARSED (tag 0x28)

	LONG	image;         // read-only in real VB6; default only
	bool	hasDC;

	bool	visible;
	char	tag[256];      // the Tag property (unrelated to this struct's own "tag bytes")

	bool	mdiChild;
	bool	keyPreview;
	bool	clipControls;
	LONG	helpContextID;

	LONG	mouseIcon;     // IPictureDisp* placeholder
	LONG	font;          // IFontDisp* placeholder -- see FontName/FontSize/... above for the scalars
	BYTE	appearance;    // 0=Flat,1=3D

	bool	whatsThisButton;  // PARSED presence only (tag 0x42)
	bool	whatsThisHelp;
	bool	showInTaskbar;
	BYTE	rightToLeft;
	BYTE	startUpPosition; // 0=Manual,1=CenterOwner,2=CenterScreen,3=Windows Default -- tag not yet confirmed
	BYTE	oleDropMode;

	LONG	palette;       // IPictureDisp* placeholder
	BYTE	paletteMode;
	bool	moveable;

	// Not part of `_Form` itself -- the .frm's separate ClientLeft/ClientTop/
	// ClientWidth/ClientHeight keys (the form's client-area rect at design time).
	// PARSED together as one unit (tag 0x35). This is what window creation actually
	// uses; Left/Top/Width/Height above are the whole-form equivalents and aren't
	// parsed yet.
	LONG	clientLeft;   // twips; LONG_MIN sentinel = "not set, let Windows place it"
	LONG	clientTop;    // twips; LONG_MIN sentinel = "not set, let Windows place it"
	LONG	clientWidth;  // twips; always a concrete value (defaulted if not found)
	LONG	clientHeight; // twips; always a concrete value (defaulted if not found)

	// Not a real _Form property -- where THIS form's own compiled property blob
	// ends (right past its ClientRect tag if one was found, otherwise right past
	// its Caption) and its placed-Controls section begins. ParseControlTemplate
	// needs to scan from here, not from the Form's raw PublicObjectDescriptor
	// address: two different forms can each have their own same-named control
	// (e.g. both having their own "Command1"), and a scan anchored on the class
	// descriptor alone has no way to tell which form a same-named match actually
	// belongs to -- confirmed live to actually pick the WRONG form's control
	// otherwise (Form2's own "Command1" rendered at Form1's Command1's
	// Left/Top/Width/Height). nullptr if the class name signature itself wasn't
	// found at all (ParseFormTemplate returned false).
	void	*pControlsAnchor;
};

/**
 * Fills every field of *pOut with the real VB6 default for a Form that sets nothing
 * explicitly. The handful this project has actually verified live: Caption =
 * pszClassName, BackColor = &H8000000F& (vbButtonFace), BorderStyle = 2 (Sizable),
 * Enabled/ControlBox/MaxButton/MinButton = True, WindowState = 0 (Normal),
 * ClientWidth/Height = a plain 4800x3600-twip window, ClientLeft/Top = "let Windows
 * place it". Everything else is the standard documented VB6 default for that
 * property (ForeColor = &H80000012&/vbWindowText, ScaleMode = 1/Twips, Appearance =
 * 1/3D, Visible/ClipControls/ShowInTaskbar/HasDC/Moveable = True, etc.) -- these
 * haven't been individually confirmed against a compiled binary the way the parsed
 * ones have, so treat them as "documented VB6 default", not "verified live".
 */
void InitFormTemplateDefaults(
	FormTemplateProps	*pOut,
	const char			*pszClassName
);

/**
 * Locates and parses a Form-derived class's compiled "property template" blob -- a
 * tag/length/value record sequence embedded in .rdata (VB6's compiled answer to a
 * .frm's `Begin VB.Form` property block; not the separate, file-based .frx format),
 * found by signature-scanning forward from the class's own PublicObjectDescriptor for
 * a second, length-prefixed copy of the class's own name (its "Name" designer
 * property) -- confirmed live across six independently-compiled test forms. No
 * compiled struct field points to this blob directly (confirmed by an exhaustive
 * xref search), so this bounded forward scan is the only way found to reach it.
 *
 * *pOut is always initialized via InitFormTemplateDefaults first, then overridden
 * field-by-field as each property's tag is actually found -- VB6 only emits a tag
 * for a property that differs from its default, so a property with no tag present
 * correctly keeps its default value. Confirmed tag bytes (from a mix of live
 * calibration and the property-ID table published at
 * https://www.vb-decompiler.org/forms_editing.htm -- cross-checked against this
 * project's own independent findings: 0x03/BackColor and 0x24/LinkTopic matched
 * exactly, which is why the rest of that page's table is trusted for the ones this
 * project hadn't already confirmed; 0x09 is the one confirmed exception -- that page
 * lists it as Visible, but a calibration form with only Enabled changed produced it,
 * so it's Enabled here, trusting this project's own controlled test over the
 * secondary source. Also note: whatever tags a form actually emits appear in
 * ascending numeric order in the blob regardless of the properties' declaration
 * order in the .frm -- confirmed by comparing a form with only MaxButton/MinButton
 * changed (emits 0x26, 0x27) against one with MaxButton/MinButton/ControlBox all
 * changed (emits 0x26, 0x27, 0x28); a first guess at this mapping had 0x26 as
 * ControlBox by assuming alphabetical order matched numeric order, which it doesn't
 * -- that produced a real, live bug (ControlBox misread as False, stripping
 * WS_SYSMENU/the close button, on a form that never touched ControlBox):
 *   0x09  Enabled          1-byte boolean (see note above re: vb-decompiler.org's Visible claim)
 *   0x0A  WindowState      1-byte value (0=Normal,1=Minimized,2=Maximized)
 *   0x0D  Caption          WORD length, text, null terminator
 *   0x03  BackColor        4-byte raw OLE_COLOR
 *   0x22  BorderStyle      1-byte value
 *   0x23  Icon             4-byte stdole.Picture reference (0xFFFFFFFF = none set)
 *   0x24  LinkTopic        WORD length, text, null terminator
 *   0x26  MaxButton        1-byte boolean
 *   0x27  MinButton        1-byte boolean
 *   0x28  ControlBox       1-byte boolean
 *   0x35  ClientRect       4 consecutive DWORDs: Left, Top, Width, Height (twips)
 *   0x42  WhatsThisButton  1-byte boolean
 * Icon and WhatsThisButton are only checked for presence (FormTemplateProps.icon
 * flips to a "has a value" state, whatsThisButton is read as a real boolean) -- Icon
 * doesn't get its actual picture data parsed (see the struct's OLE-placeholder note).
 * Every other _Form member documented in FormTemplateProps but not listed above
 * simply isn't parsed yet; add its tag here (and to the FindTag scan in the .cpp)
 * once confirmed the same way these were -- compile a calibration form with only
 * that one property changed and diff the blob.
 *
 * Returns false (with *pOut left at pure defaults) if even the class's own name
 * signature isn't found -- callers can treat that the same as "nothing overridden"
 * rather than as fatal, since *pOut is already fully valid either way.
 */
bool ParseFormTemplate(
	ObjectInfoWithOptional	*pDesc,
	FormTemplateProps		*pOut
);
