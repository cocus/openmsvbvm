#pragma once

#include "vba_internal.h"
#include "vba_structures.h"

/**
 * Compiled-in designer properties for one control placed on a Form, read straight
 * out of the compiled EXE's own data (twips, matching the .frm source) -- the
 * control-level equivalent of FormProperties.hpp's FormTemplateProps.
 *
 * Only the common base properties every control shares are modeled here (this
 * project only supports CommandButton today -- see ControlWindow.hpp); type-specific
 * properties would need their own confirmed tags/offsets added the same way these
 * were, once a second control type is supported.
 */
struct ControlTemplateProps
{
	char	name[64];      // Name -- the control's own name, e.g. "Command1"
	char	caption[256];  // Caption (or Text, depending on control type) -- PARSED

	// twips, matching the .frm source's own Left/Top/Width/Height keys -- PARSED
	LONG	left;
	LONG	top;
	LONG	width;
	LONG	height;
};

void InitControlTemplateDefaults(ControlTemplateProps *pOut, const char *pszControlName);

/**
 * Locates and parses a placed control's compiled property record -- confirmed live
 * against two independently-compiled calibration builds (one where Caption happened
 * to equal Name, one where it deliberately didn't, to rule out a name-matching scan
 * being a coincidence) by signature-scanning FORWARD ONLY from pControlsAnchor (see
 * FormTemplateProps.pControlsAnchor, FormProperties.hpp -- the exact position right
 * past the owning Form's own property blob, NOT the Form's raw PublicObjectDescriptor
 * address) for a length-prefixed copy of the control's own name, then parsing the
 * fixed-shape record that follows it byte-for-byte:
 *   WORD nameLen, name text, 0x00, 2 unknown bytes,
 *   WORD captionLen, caption text, 0x00, 1 unknown byte,
 *   WORD Left, WORD Top, WORD Width, WORD Height, ... (remainder not parsed yet)
 *
 * Anchoring on the Form's raw descriptor and scanning both directions (like
 * ParseFormTemplate now does for the Form's own blob) was tried first and confirmed
 * BROKEN: two different forms can each place their own same-named control (e.g. both
 * having their own "Command1"), and nothing about a name-only match anchored that
 * loosely can tell which owning form it actually belongs to -- confirmed live to
 * silently pick the WRONG form's control (Form2's own Command1 rendered at Form1's
 * Command1's Left/Top/Width/Height). Anchoring tightly on pControlsAnchor instead
 * -- confirmed to sit only a few bytes before the real match in every case checked
 * -- removes the ambiguity entirely: only forward, only within this owning form's
 * own Controls section.
 *
 * *pOut is always initialized via InitControlTemplateDefaults first. Returns false
 * (with *pOut left at pure defaults) if pControlsAnchor is null or the control's name
 * signature isn't found.
 */
bool ParseControlTemplate(void *pControlsAnchor, const char *pszControlName, ControlTemplateProps *pOut);
