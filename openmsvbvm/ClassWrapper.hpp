#pragma once

#include "vba_internal.h"
#include "vba_structures.h"
#include "ObjectWrapper.hpp"

/**
 * Plain VB6 class-module instance -- mirrors real msvbvm60's `CClassModule` (confirmed
 * via IDA on the symbol-rich real DLL: its constructor sets its own vtables directly,
 * with no base-class constructor call at all -- a standalone class, unrelated to
 * `CTL`/`FORM`). This project's own `vbControlWrapper` (ControlWrapper.hpp, the
 * `CTL` equivalent Forms and any future intrinsic control derive from) is a
 * deliberately separate, sibling lineage -- neither derives from the other, matching
 * that real independence at the level that matters: a Form isn't "a plain class plus
 * a window". Both happen to derive from vbObjectWrapper (ObjectWrapper.hpp) for this
 * project's own DRY reasons -- see that class's header comment.
 *
 * Adds nothing over vbObjectWrapper today; exists as its own type so __vbaNew
 * (ObjectManipulation.cpp) constructs a type genuinely distinct from vbControlWrapper/
 * vbFormWrapper for a plain class, and so a dynamic_cast<vbClassWrapper*> would
 * correctly reject a Form the same way dynamic_cast<vbFormWrapper*> already rejects a
 * plain class.
 */
class vbClassWrapper : public vbObjectWrapper
{
public:
	vbClassWrapper(vba_VBVTable * pWrapperVtable, ObjectInfoWithOptional* pObjInfo)
		: vbObjectWrapper(pWrapperVtable, pObjInfo)
	{
	}
};
