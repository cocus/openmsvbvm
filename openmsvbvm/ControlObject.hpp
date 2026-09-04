#pragma once

#include "vba_internal.h"

/**
 * Constructs a real, standalone "control accessor" COM object backing a placed
 * control's own HWND -- what a Form's own compiled code gets back when it accesses a
 * placed control by name (e.g. "Label1.Caption = ..."). See ObjectManipulation.cpp's
 * __vbaNew and FormWrapper.cpp's vbFormWrapper::GetControlAccessor for how/when this
 * gets called, and Label.idl's header comment for the confirmed real vtable layout
 * this implements.
 *
 * Implements _Label's Name/Caption for real (against hwndControl directly); every
 * other _Label member is a stub. Reused as-is for ANY placed control type this
 * project renders (CommandButton included), not just a real Label -- their early
 * Name/Caption slots are confirmed identical (see Label.idl's header comment), and
 * this project doesn't have a dedicated _CommandButton.idl yet. hwndControl may be
 * null (an unsupported/unrendered control type was still referenced by name in
 * code) -- Name still works in that case, Caption reads back empty and silently
 * no-ops on write.
 *
 * Returns a new object with a refcount of 1 (the caller owns this reference).
 */
IDispatch * CreateControlAccessorObject(HWND hwndControl, const char *pszControlName);
