#include "vba_internal.h"
#include "Logging.hpp"
#include "ControlWindow.hpp"

extern HINSTANCE g_dllModuleHandle;

/* Shared by every CreateXxxControl below -- twips -> pixels at the real screen DPI,
   same conversion FormWindow.cpp's CreateFormWindow already uses. */
static void ControlRectFromTwips(const ControlWindowCreateParams& params, int& x, int& y, int& cx, int& cy)
{
    HDC hdc = GetDC(nullptr);
    int dpiX = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
    int dpiY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
    if (hdc)
    {
        ReleaseDC(nullptr, hdc);
    }

    x = MulDiv(params.leftTwips, dpiX, 1440);
    y = MulDiv(params.topTwips, dpiY, 1440);
    cx = MulDiv(params.widthTwips, dpiX, 1440);
    cy = MulDiv(params.heightTwips, dpiY, 1440);
}

HWND CreateButtonControl(HWND hwndParent, const ControlWindowCreateParams& params)
{
    if (!hwndParent)
    {
        return nullptr;
    }

    int x, y, cx, cy;
    ControlRectFromTwips(params, x, y, cx, cy);

    return CreateWindowExW(
        0,
        L"BUTTON",
        params.caption,
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        x,
        y,
        cx,
        cy,
        hwndParent,
        (HMENU)(INT_PTR)params.controlId,
        g_dllModuleHandle,
        nullptr);
}

HWND CreateLabelControl(HWND hwndParent, const ControlWindowCreateParams& params)
{
    if (!hwndParent)
    {
        return nullptr;
    }

    int x, y, cx, cy;
    ControlRectFromTwips(params, x, y, cx, cy);

    return CreateWindowExW(
        0,
        L"STATIC",
        params.caption,
        WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOTIFY,
        x,
        y,
        cx,
        cy,
        hwndParent,
        (HMENU)(INT_PTR)params.controlId,
        g_dllModuleHandle,
        nullptr);
}
