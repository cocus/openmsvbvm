#include "vba_internal.h"
#include "Logging.hpp"
#include "FormProperties.hpp"

#include <cstring>
#include <climits>

/* How far from PublicObjectDescriptor to scan (in EITHER direction) for the class's
   own length-prefixed name (confirmed live: real distances seen were ~0x8F6 to
   ~0xACE bytes AHEAD for tiny single-form calibration projects; a multi-class test
   project (Module1 + 3 class modules + Form1) instead put Form1's template blob
   ~0xEB8 bytes BEHIND its PublicObjectDescriptor -- confirmed live by dumping
   Proyecto1.exe's raw bytes and finding the real, tag-bearing blob sitting at a
   LOWER address than the descriptor once other objects' own compiled code/data came
   between them. So which side the blob ends up on depends on how many other
   objects the compiler emitted around this one -- scan both directions rather than
   assume forward-only). No compiled pointer reaches this blob (verified via an
   exhaustive xref search over the whole region in a compiled test EXE), so a bounded
   scan from a known-good anchor is the only way found to locate it. */
static const size_t kNameScanWindow = 0x4000;

/* How far past Caption's end to look for the optional tags below -- confirmed live
   to all sit within a few dozen bytes of it. Kept small since single-byte tags are
   more likely to coincidentally collide with unrelated bytes the further out the
   scan goes. */
static const size_t kTagScanWindow = 0x40;

void InitFormTemplateDefaults(FormTemplateProps* pOut, const char* pszClassName)
{
    memset(pOut, 0, sizeof(*pOut));

    strncpy_s(pOut->name, sizeof(pOut->name), pszClassName ? pszClassName : "Form", _TRUNCATE);
    strncpy_s(pOut->caption, sizeof(pOut->caption), pszClassName ? pszClassName : "Form", _TRUNCATE);
    strncpy_s(pOut->linkTopic, sizeof(pOut->linkTopic), pszClassName ? pszClassName : "Form", _TRUNCATE);

    // -- Verified live (see ParseFormTemplate's tag table) --
    pOut->backColor = 0x8000000F; // vbButtonFace
    pOut->borderStyle = 2;        // Sizable
    pOut->enabled = true;
    pOut->controlBox = true;
    pOut->maxButton = true;
    pOut->minButton = true;
    pOut->windowState = 0; // Normal
    pOut->icon = -1;       // "has the default VB icon", matching the 0xFFFFFFFF sentinel seen live
    pOut->whatsThisButton = false;
    pOut->clientLeft = LONG_MIN;
    pOut->clientTop = LONG_MIN;
    pOut->clientWidth = 4800;
    pOut->clientHeight = 3600;

    // -- Documented VB6 defaults, not yet individually confirmed against a compiled
    //    binary (see the struct's doc comment) --
    pOut->foreColor = 0x80000012; // vbWindowText
    pOut->left = LONG_MIN;
    pOut->top = LONG_MIN;
    pOut->width = 4950;
    pOut->height = 3690;
    pOut->mousePointer = 0; // Default
    strncpy_s(pOut->fontName, sizeof(pOut->fontName), "MS Sans Serif", _TRUNCATE);
    pOut->fontSize = 8.25f;
    pOut->fontBold = false;
    pOut->fontItalic = false;
    pOut->fontStrikethru = false;
    pOut->fontUnderline = false;
    pOut->fontTransparent = true;
    pOut->scaleLeft = 0;
    pOut->scaleTop = 0;
    pOut->scaleWidth = pOut->clientWidth;
    pOut->scaleHeight = pOut->clientHeight;
    pOut->scaleMode = 1; // Twips
    pOut->drawStyle = 0; // Solid
    pOut->drawWidth = 1;
    pOut->fillStyle = 1; // Transparent
    pOut->fillColor = 0x00000000;
    pOut->drawMode = 13; // Copy Pen
    pOut->autoRedraw = false;
    pOut->picture = 0;  // none
    pOut->linkMode = 0; // None
    pOut->image = 0;
    pOut->hasDC = true;
    pOut->visible = true;
    pOut->tag[0] = '\0';
    pOut->mdiChild = false;
    pOut->keyPreview = false;
    pOut->clipControls = true;
    pOut->helpContextID = 0;
    pOut->mouseIcon = 0;
    pOut->font = 0;
    pOut->appearance = 1; // 3D
    pOut->whatsThisHelp = false;
    pOut->showInTaskbar = true;
    pOut->rightToLeft = 0;
    pOut->startUpPosition = 3; // Windows Default
    pOut->oleDropMode = 0;
    pOut->palette = 0;
    pOut->paletteMode = 0; // Halftone
    pOut->moveable = true;
}

/* Every tag this parser looks for past Caption is a single byte; a hit is only
   accepted once its value also looks plausible for that specific property (a single
   byte tag has real collision risk against unrelated data otherwise), so these keep
   scanning past an implausible match instead of stopping at the first occurrence. */

/* BorderStyle (0-5) / WindowState (0-2) -- any small-range single-byte enum. */
static bool FindEnumTag(BYTE* pStart, BYTE* pEnd, BYTE tag, BYTE maxValid, BYTE* pOutValue)
{
    for (BYTE* p = pStart; p + 1 < pEnd; p++)
    {
        if (*p == tag && p[1] <= maxValid)
        {
            *pOutValue = p[1];
            return true;
        }
    }
    return false;
}

/* Enabled / ControlBox / MaxButton / MinButton -- VB6 booleans seen compiled as 0x00
   or 0x01 in this project's own calibration; also accept the classic VB 0xFF "True"
   encoding in case a differently-compiled project uses it. */
static bool FindBoolTag(BYTE* pStart, BYTE* pEnd, BYTE tag, bool* pOutValue)
{
    for (BYTE* p = pStart; p + 1 < pEnd; p++)
    {
        if (*p == tag && (p[1] == 0x00 || p[1] == 0x01 || p[1] == 0xFF))
        {
            *pOutValue = (p[1] != 0x00);
            return true;
        }
    }
    return false;
}

/* BackColor -- a 4-byte OLE_COLOR has no small fixed value range like the enum/bool
   tags above, so this checks the encoding's own shape instead: either a plain
   0x00BBGGRR value (top byte 0) or a system-color reference (top byte 0x80, the two
   bytes below it always 0, low byte a COLOR_* index). Without this check a single
   coincidental 0x03 byte anywhere in the scan window would be believed as a real
   BackColor tag and the 4 bytes after it taken as the color verbatim -- confirmed
   live to actually happen (a form with no BackColor override at all still "found" a
   tag, reading unrelated bytes as a color that rendered as solid red). */
static bool FindColorTag(BYTE* pStart, BYTE* pEnd, BYTE tag, DWORD* pOutValue)
{
    for (BYTE* p = pStart; p + 5 <= pEnd; p++)
    {
        if (*p != tag)
        {
            continue;
        }

        DWORD value = *(DWORD*)(p + 1);
        bool bPlausible = ((value >> 24) == 0x00) || ((value >> 24) == 0x80 && (value & 0x00FFFF00) == 0);
        if (!bPlausible)
        {
            continue;
        }

        *pOutValue = value;
        return true;
    }
    return false;
}

bool ParseFormTemplate(ObjectInfoWithOptional* pDesc, FormTemplateProps* pOut)
{
    const char* pszName = (pDesc && pDesc->hdr.lpObject) ? pDesc->hdr.lpObject->lpszObjectName : nullptr;
    InitFormTemplateDefaults(pOut, pszName);

    if (!pDesc || !pszName)
    {
        return false;
    }

    size_t nameLen = strlen(pszName);
    if (nameLen == 0 || nameLen > 200)
    {
        return false;
    }

    bool bFound = false;

    __try
    {
        BYTE* pAnchor = (BYTE*)pDesc->hdr.lpObject;
        BYTE* pScanStart = pAnchor - kNameScanWindow;
        BYTE* pScanEnd = pAnchor + kNameScanWindow;

        // Clamp the backward edge to the start of whatever VirtualAlloc'd region the
        // anchor itself lives in (for a loaded module, that's the module's own load
        // base) -- pAnchor sits close enough to the start of a small EXE's image that
        // a plain "pAnchor - kNameScanWindow" can land before the image is even
        // mapped, which would fault on the very first read and, since that's the
        // FIRST byte the loop touches, abort this whole __try before it ever reaches
        // the real (in-bounds) part of the window.
        MEMORY_BASIC_INFORMATION mbi;
        if (VirtualQuery(pAnchor, &mbi, sizeof(mbi)) && (BYTE*)mbi.AllocationBase > pScanStart)
        {
            pScanStart = (BYTE*)mbi.AllocationBase;
        }

        BYTE* pFound = nullptr;
        for (BYTE* p = pScanStart; p + 2 + nameLen + 1 < pScanEnd; p++)
        {
            if (*(WORD*)p != (WORD)nameLen)
            {
                continue;
            }
            if (memcmp(p + 2, pszName, nameLen) != 0)
            {
                continue;
            }
            if (p[2 + nameLen] != 0)
            {
                continue;
            }
            pFound = p;
            break;
        }

        if (!pFound)
        {
            return false;
        }

        bFound = true;

        /* Caption: 1-byte tag 0x0D, WORD length, text, null. */
        BYTE* pCursor = pFound + 2 + nameLen + 1;
        if (*pCursor == 0x0D)
        {
            pCursor += 2; // tag byte + the "differs from default" flag byte seen after it
            WORD captionLen = *(WORD*)pCursor;
            pCursor += 2;

            if (captionLen < sizeof(pOut->caption))
            {
                memcpy(pOut->caption, pCursor, captionLen);
                pOut->caption[captionLen] = '\0';
            }
            pCursor += captionLen + 1;
        }

        /* Every property below is only present at all when set to a non-default
           value -- scan for each tag independently rather than assuming a fixed
           position/order (see the header's tag table for the sources). */
        BYTE* pTagScanEnd = pCursor + kTagScanWindow;

        // ClientRect (tag 0x35) is searched for FIRST, ahead of every other tag below
        // it in ID order, because its own end is also the real upper bound for all of
        // them: right past a form's own last tag comes the unrelated Controls section
        // (placed-control names/captions/etc, ControlProperties.cpp), and a stray
        // byte in THAT data can otherwise be misread as one of the single-byte tags
        // below -- confirmed live (a form with no BackColor override at all still
        // "found" tag 0x03 a few bytes past its own ClientRect, its 4 "value" bytes
        // coming from the next placed control's own record, rendering as solid red).
        // Since tags are confirmed to appear in ascending numeric order, every tag ID
        // below 0x35 that's actually present must sit BEFORE this point, so clamping
        // their own scans to end here (instead of the full kTagScanWindow) costs
        // nothing real while closing off that whole class of false positive.
        // WhatsThisButton (0x42) is the one confirmed tag ABOVE 0x35, so it alone
        // still scans out to the full window.
        BYTE* pBeforeClientRect = pTagScanEnd;
        for (BYTE* p = pCursor; p + 17 <= pTagScanEnd; p++)
        {
            if (*p != 0x35)
            {
                continue;
            }

            LONG* pGeom = (LONG*)(p + 1);
            bool bPlausible = pGeom[2] > 0 && pGeom[2] < 500000 && pGeom[3] > 0 && pGeom[3] < 500000 && pGeom[0] > -500000 &&
                              pGeom[0] < 500000 && pGeom[1] > -500000 && pGeom[1] < 500000;

            if (bPlausible)
            {
                pOut->clientLeft = pGeom[0];
                pOut->clientTop = pGeom[1];
                pOut->clientWidth = pGeom[2];
                pOut->clientHeight = pGeom[3];
                pBeforeClientRect = p;
                break;
            }
        }

        // Right past ClientRect's own 17 bytes is where the Controls section starts
        // for real (confirmed live -- see the comment on FormTemplateProps
        // .pControlsAnchor); when no ClientRect was found at all, pCursor (right
        // past Caption) is the best fallback anchor available.
        pOut->pControlsAnchor = (pBeforeClientRect != pTagScanEnd) ? (pBeforeClientRect + 17) : pCursor;

        BYTE byteValue;

        FindBoolTag(pCursor, pBeforeClientRect, 0x09, &pOut->enabled);
        if (FindEnumTag(pCursor, pBeforeClientRect, 0x0A, 2, &byteValue))
            pOut->windowState = byteValue;
        if (FindEnumTag(pCursor, pBeforeClientRect, 0x22, 5, &byteValue))
            pOut->borderStyle = byteValue;
        // Tags are emitted in ascending numeric order regardless of declaration
        // order (confirmed live: a form with only MaxButton/MinButton changed
        // produced tags 0x26/0x27, not 0x27/0x28 as a naive "alphabetical ==
        // numeric order" guess first assumed) -- 0x26=MaxButton, 0x27=MinButton,
        // 0x28=ControlBox.
        FindBoolTag(pCursor, pBeforeClientRect, 0x26, &pOut->maxButton);
        FindBoolTag(pCursor, pBeforeClientRect, 0x27, &pOut->minButton);
        FindBoolTag(pCursor, pBeforeClientRect, 0x28, &pOut->controlBox);
        FindBoolTag(pCursor, pTagScanEnd, 0x42, &pOut->whatsThisButton);

        FindColorTag(pCursor, pBeforeClientRect, 0x03, &pOut->backColor);

        // Icon: presence only -- the raw stdole.Picture reference is kept verbatim
        // (see the struct's OLE-placeholder note), not parsed into real picture data.
        for (BYTE* p = pCursor; p + 5 <= pBeforeClientRect; p++)
        {
            if (*p == 0x23)
            {
                pOut->icon = *(LONG*)(p + 1);
                break;
            }
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        InitFormTemplateDefaults(pOut, pszName);
        return false;
    }

    return bFound;
}
