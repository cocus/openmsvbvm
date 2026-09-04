#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "ObjectWrapper.hpp"

/**
 * Base for every "visual/intrinsic" VB6 object -- mirrors real msvbvm60's `CTL`
 * (confirmed via IDA on the symbol-rich real DLL: `FORM::FORM(XMOD*, DESK*)` calls
 * `CTL::CTL(XMOD*, DESK*)` first, then overrides only 2 of CTL's 7 interfaces; every
 * other intrinsic control -- CHECK, COMBO, TIMER, MDI, APP, SCRN, ~25 more -- derives
 * from CTL the same way). vbFormWrapper (FormWrapper.hpp) derives from this instead
 * of from vbClassWrapper -- a Form is conceptually "a control-like visual object",
 * not "a plain class module plus a window". Any future intrinsic object this project
 * gives a real window/visual presence to (MDIForm as its own thing, an actual
 * control) belongs under this class too, not under vbClassWrapper.
 *
 * Adds nothing over vbObjectWrapper today; exists as its own type purely for that
 * taxonomy -- see vbObjectWrapper's header comment for why the actual
 * IDispatch/IConnectionPointContainer bridge implementation is shared with
 * vbClassWrapper via vbObjectWrapper rather than duplicated the way CTL/CClassModule
 * really are in msvbvm60.
 */
class vbControlWrapper : public vbObjectWrapper
{
public:
	vbControlWrapper(vba_VBVTable * pWrapperVtable, ObjectInfoWithOptional* pObjInfo)
		: vbObjectWrapper(pWrapperVtable, pObjInfo)
	{
	}
};
