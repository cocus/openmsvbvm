#include "vba_internal.h"
#include "Logging.hpp"
#include "ControlProperties.hpp"

#include <cstring>
#include <climits>

/* No compiled pointer reaches a control's property record directly,
   seems to sit only a few bytes past pControlsAnchor (see
   FormTemplateProps.pControlsAnchor, FormProperties.hpp, the exact position right
   past the owning Form's own compiled property blob), so this only has to scan
   FORWARD from there, unlike FormProperties.cpp's own bidirectional
   kNameScanWindow scan for the Form's own blob (which has no such precise anchor to
   start from). Kept generously large anyway as a safety margin, not because the
   real distance needs it. */
static const size_t kNameScanWindow = 0x4000;

void InitControlTemplateDefaults(ControlTemplateProps* pOut, const char* pszControlName)
{
    memset(pOut, 0, sizeof(*pOut));

    strncpy_s(pOut->name, sizeof(pOut->name), pszControlName ? pszControlName : "", _TRUNCATE);
    strncpy_s(pOut->caption, sizeof(pOut->caption), pszControlName ? pszControlName : "", _TRUNCATE);

    // Documented VB6 default CommandButton size/position -- not individually
    // confirmed against a compiled binary (a real placed control's Left/Top/Width/
    // Height are always found in practice, so this default is mostly theoretical).
    pOut->left = 0;
    pOut->top = 0;
    pOut->width = 1215;
    pOut->height = 495;
}

bool ParseControlTemplate(void* pControlsAnchor, const char* pszControlName, ControlTemplateProps* pOut)
{
    InitControlTemplateDefaults(pOut, pszControlName);

    if (!pControlsAnchor || !pszControlName)
    {
        return false;
    }

    size_t nameLen = strlen(pszControlName);
    if (nameLen == 0 || nameLen > 200)
    {
        return false;
    }

    __try
    {
        BYTE* pScanStart = (BYTE*)pControlsAnchor;
        BYTE* pScanEnd = pScanStart + kNameScanWindow;

        BYTE* pFound = nullptr;
        for (BYTE* p = pScanStart; p + 2 + nameLen + 1 < pScanEnd; p++)
        {
            if (*(WORD*)p != (WORD)nameLen)
            {
                continue;
            }
            if (memcmp(p + 2, pszControlName, nameLen) != 0)
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

        /* Name: WORD length, text, null terminator, then 2 unrecognized bytes
           (seen live as 04 01 in every sample) before Caption's own record starts. */
        BYTE* pCursor = pFound + 2 + nameLen + 1 + 2;

        WORD captionLen = *(WORD*)pCursor;
        pCursor += 2;

        if (captionLen < sizeof(pOut->caption))
        {
            memcpy(pOut->caption, pCursor, captionLen);
            pOut->caption[captionLen] = '\0';
        }
        pCursor += captionLen + 1; // text + null terminator

        /* 1 more unrecognized byte (seen live as 04) before the Left/Top/Width/Height
           quad -- confirmed live across two independently-compiled calibration builds
           with different Caption/Left/Top/Width/Height values (Caption deliberately
           NOT matching Name in the second, to rule out this scan just re-finding the
           Name string by coincidence). */
        pCursor += 1;

        LONG left = *(WORD*)pCursor;
        pCursor += 2;
        LONG top = *(WORD*)pCursor;
        pCursor += 2;
        LONG width = *(WORD*)pCursor;
        pCursor += 2;
        LONG height = *(WORD*)pCursor;
        pCursor += 2;

        // Plausibility check (same spirit as FormProperties.cpp's own ClientRect
        // check) before trusting these over the defaults.
        bool bPlausible = width > 0 && width < 500000 && height > 0 && height < 500000 && left > -500000 && left < 500000 &&
                          top > -500000 && top < 500000;

        if (bPlausible)
        {
            pOut->left = left;
            pOut->top = top;
            pOut->width = width;
            pOut->height = height;
        }

        return true;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        InitControlTemplateDefaults(pOut, pszControlName);
        return false;
    }
}
