#include "vba_internal.h"
#include "Logging.hpp"
#include "FormWindow.hpp"
#include "ObjectManipulation.hpp"

#include <climits>

extern HINSTANCE g_dllModuleHandle;

static const wchar_t* kOwnerPropName = L"OpenMsvbvmOwner";
static const wchar_t* kBackBrushPropName = L"OpenMsvbvmBackBrush";

/* OLE_COLOR -> COLORREF: a system-color reference has 0x80 in the top byte and a
   COLOR_* index in the low byte; anything else is already a plain 0x00BBGGRR value,
   the same byte order Windows' own COLORREF uses. */
static COLORREF OleColorToColorRef(DWORD oleColor)
{
    if ((oleColor & 0xFF000000) == 0x80000000)
    {
        return GetSysColor(oleColor & 0xFF);
    }

    return (COLORREF)(oleColor & 0xFFFFFF);
}

/* VB6 BorderStyle -> window style/ex-style. Not pixel-exact to real VB6 (e.g. Fixed
   Single still normally shows minimize/maximize boxes, kept here for simplicity) --
   the main distinctions (resizable vs fixed frame, bordered vs none, tool-window
   caption) are what's implemented. ControlBox/MaxButton/MinButton (confirmed tags
   0x26/0x27/0x28 -- see FormProperties.hpp) then strip the corresponding style bits
   individually, matching real VB6 semantics where those are independent of
   BorderStyle (e.g. a Sizable form can still turn off just the maximize box). */
static void BorderStyleToWindowStyle(BYTE borderStyle, bool controlBox, bool maxButton, bool minButton, DWORD& style, DWORD& exStyle)
{
    style = WS_OVERLAPPEDWINDOW;
    exStyle = 0;

    switch (borderStyle)
    {
    case 0: // None
        style = WS_POPUP;
        break;
    case 1: // Fixed Single
        style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;
        break;
    case 3: // Fixed Dialog
        style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU;
        break;
    case 4: // Fixed ToolWindow
        style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU;
        exStyle = WS_EX_TOOLWINDOW;
        break;
    case 5: // Sizable ToolWindow
        style = WS_OVERLAPPEDWINDOW;
        exStyle = WS_EX_TOOLWINDOW;
        break;
    default: // 2 = Sizable (also the fallback when BorderStyle wasn't found/parsed)
        break;
    }

    if (!controlBox)
    {
        style &= ~(WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
    }
    if (!maxButton)
    {
        style &= ~WS_MAXIMIZEBOX;
    }
    if (!minButton)
    {
        style &= ~WS_MINIMIZEBOX;
    }
}

/* Real msvbvm60 window class name for a Form window -- confirmed via disassembly of the
   real system msvbvm60.dll: a shared helper (sub_6601760A) builds it at runtime as
   "ThunderRT6" + a per-object-kind suffix (a literal "DFrame", or one supplied by the
   caller, e.g. "Form") + optionally "DC" appended. The real, system-installed
   msvbvm60.dll's ThunderRT6Main hidden main-window class name is a separate, fully
   literal string; this project's own debug logging already echoes that convention
   ("ThunRTMain: ..."). */
static const wchar_t* kFormWindowClassName = L"ThunderRT6FormDC";
static bool g_bFormWindowClassRegistered = false;

/**
 * Shared window procedure for every VB6 Form window this runtime creates. On
 * WM_CLOSE, attempts Form_QueryUnload dispatch first (see VBFormTryQueryUnload /
 * vbFormWrapper::TryFireQueryUnload) -- if the handler sets Cancel, the close is
 * aborted, matching real VB6 semantics. On WM_COMMAND (a placed control's
 * notification, e.g. a CommandButton's BN_CLICKED), forwards to VBFormHandleCommand /
 * vbFormWrapper::HandleCommand. Other intrinsic events (Unload, Resize, ...) aren't
 * dispatched yet.
 */
static LRESULT CALLBACK FormWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CLOSE:
    {
        void* pOwner = (void*)GetPropW(hwnd, kOwnerPropName);
        short cancel = 0;
        if (pOwner)
        {
            VBFormTryQueryUnload(pOwner, &cancel);
        }
        if (!cancel)
        {
            DestroyWindow(hwnd);
        }
        return 0;
    }

    case WM_COMMAND:
    {
        void* pOwner = (void*)GetPropW(hwnd, kOwnerPropName);
        if (pOwner)
        {
            VBFormHandleCommand(pOwner, LOWORD(wParam), HIWORD(wParam));
        }
        return 0;
    }

    case WM_ERASEBKGND:
    {
        HBRUSH hBrush = (HBRUSH)GetPropW(hwnd, kBackBrushPropName);
        if (hBrush)
        {
            RECT rc;
            GetClientRect(hwnd, &rc);
            FillRect((HDC)wParam, &rc, hBrush);
            return 1;
        }
        break;
    }

    case WM_DESTROY:
    {
        HWND* pHwndSlot = (HWND*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
        if (pHwndSlot)
        {
            *pHwndSlot = nullptr;
        }
        RemovePropW(hwnd, kOwnerPropName);

        HBRUSH hBrush = (HBRUSH)GetPropW(hwnd, kBackBrushPropName);
        if (hBrush)
        {
            DeleteObject(hBrush);
            RemovePropW(hwnd, kBackBrushPropName);
        }
        return 0;
    }
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void EnsureFormWindowClassRegistered()
{
    if (g_bFormWindowClassRegistered)
    {
        return;
    }

    WNDCLASSEXW wc = {sizeof(wc)};
    wc.style = CS_DBLCLKS; // VB6 forms deliver DblClick events; needs this class style
    wc.lpfnWndProc = FormWndProc;
    wc.hInstance = g_dllModuleHandle;
    wc.hCursor = LoadCursorW(nullptr, MAKEINTRESOURCEW(32512) /* IDC_ARROW */);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = kFormWindowClassName;

    RegisterClassExW(&wc);
    g_bFormWindowClassRegistered = true;
}

HWND CreateFormWindow(HWND* pHwndSlot, const FormWindowCreateParams& params)
{
    EnsureFormWindowClassRegistered();

    DWORD style, exStyle;
    BorderStyleToWindowStyle(params.borderStyle, params.controlBox, params.maxButton, params.minButton, style, exStyle);

    /* Twips -> pixels at the real screen DPI (1440 twips/inch), then grow the
       desired CLIENT size out to a window size for this BorderStyle's
       borders/title bar via AdjustWindowRectEx, matching what ClientWidth/
       ClientHeight actually describe in the .frm source. Width/height are always
       concrete (FormProperties.hpp defaults them); position uses CW_USEDEFAULT
       when left/top weren't found (LONG_MIN sentinel). */
    HDC hdc = GetDC(nullptr);
    int dpiX = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
    int dpiY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
    if (hdc)
    {
        ReleaseDC(nullptr, hdc);
    }

    int x = (params.clientLeftTwips == LONG_MIN) ? CW_USEDEFAULT : MulDiv(params.clientLeftTwips, dpiX, 1440);
    int y = (params.clientTopTwips == LONG_MIN) ? CW_USEDEFAULT : MulDiv(params.clientTopTwips, dpiY, 1440);

    RECT rc = {0, 0, MulDiv(params.clientWidthTwips, dpiX, 1440), MulDiv(params.clientHeightTwips, dpiY, 1440)};
    AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    int cx = rc.right - rc.left;
    int cy = rc.bottom - rc.top;

    HWND hwnd = CreateWindowExW(
        exStyle, kFormWindowClassName, params.caption, style, x, y, cx, cy, nullptr, nullptr, g_dllModuleHandle, nullptr);

    if (hwnd)
    {
        *pHwndSlot = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)pHwndSlot);
        SetPropW(hwnd, kOwnerPropName, (HANDLE)params.pOwnerVBVTable);

        HBRUSH hBrush = CreateSolidBrush(OleColorToColorRef(params.backColor));
        if (hBrush)
        {
            SetPropW(hwnd, kBackBrushPropName, (HANDLE)hBrush);
        }

        if (!params.enabled)
        {
            EnableWindow(hwnd, FALSE);
        }
    }

    return hwnd;
}

void RunModalMessageLoop(HWND* pHwndSlot)
{
    MSG msg;
    while (*pHwndSlot && GetMessageW(&msg, nullptr, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}
