#include "vba_internal.h"
#include "vba_exception.h"
#include "resource.h"
#include <Windows.h>

#include "vba_strManipulation.h"

typedef struct
{
	LONG			lX;
	bool			bHasX;
	LONG			lY;
	bool			bHasY;
	wchar_t			*lpwstrTitle;
	wchar_t			*lpwstrPrompt;
	wchar_t			*lpwstrDefaultValue;
	wchar_t			*lpwstrHelpFile;
	wchar_t			*lpwstrHelpContext;
	wchar_t			*lpwstrReturnValue;

} inputBoxStorage_t;

static BOOL CALLBACK InputBoxDlgProc(
	HWND		hwnd,
	UINT		message,
	WPARAM		wParam,
	LPARAM		lParam
)
{
	switch (message)
	{
		case WM_INITDIALOG:
		{
			/* Code from https://docs.microsoft.com/en-us/windows/desktop/dlgbox/using-dialog-boxes */
			// Get the owner window and dialog box rectangles. 
			HWND		hwndParent;
			RECT		rcDlg;
			RECT		rcOwner;
			RECT		rc;

			if ((hwndParent = GetParent(hwnd)) == NULL)
			{
				hwndParent = GetDesktopWindow();
			}

			GetWindowRect(hwndParent, &rcOwner);
			GetWindowRect(hwnd, &rcDlg);
			CopyRect(&rc, &rcOwner);

			// Offset the owner and dialog box rectangles so that right and bottom 
			// values represent the width and height, and then offset the owner again 
			// to discard space taken up by the dialog box. 
			OffsetRect(&rcDlg, -rcDlg.left, -rcDlg.top);
			OffsetRect(&rc, -rc.left, -rc.top);
			OffsetRect(&rc, -rcDlg.right, -rcDlg.bottom);

			/* Setup the pointer to the inputBoxStorage for later usage */
			SetWindowLongPtr(hwnd, DWLP_USER, lParam);

			/* Retrieve the storage structure from the user defined parameter */
			inputBoxStorage_t * ibs = (inputBoxStorage_t*)lParam;
			if (!ibs)
			{
				return FALSE;
			}

			if (!ibs->bHasX)
			{
				ibs->lX = rcOwner.left + (rc.right / 2);
			}
			if (!ibs->bHasY)
			{
				ibs->lY = rcOwner.top + (rc.bottom / 2);
			}

			/* Set the window pos with the specified X, Y values */
			SetWindowPos(
				hwnd,
				HWND_TOP,
				ibs->lX,
				ibs->lY,
				0, 0,          // Ignores size arguments. 
				SWP_NOSIZE
			);

			HWND h = NULL;


			/* If there's a helpfile present, show the Help button */
			if (ibs->lpwstrHelpFile)
			{
				h = GetDlgItem(hwnd, IDC_BUTTON_HELP);
				ShowWindow(h, SW_SHOW);
			}

			/* Set the default value (if there's one) */
			if (ibs->lpwstrDefaultValue)
			{
				h = GetDlgItem(hwnd, IDC_INPUT);
				SetWindowTextW(h, ibs->lpwstrDefaultValue);
			}

			/* Set the lpwstrPrompt (static text) if there's one */
			if (ibs->lpwstrPrompt)
			{
				h = GetDlgItem(hwnd, IDC_USER_STATIC);
				SetWindowTextW(h, ibs->lpwstrPrompt);
			}

			/* Set the window lpwstrTitle (if there's one) */
			if (ibs->lpwstrTitle)
			{
				SetWindowTextW(hwnd, ibs->lpwstrTitle);
			}

			/* Set the caption of the controls using system-dependant strings (if available) */
			typedef LPCWSTR(WINAPI *pmb_getstring)(int strId);
			HMODULE hLibUser32 = LoadLibraryA("user32.dll");
			if (hLibUser32)
			{
				/* Use the undoc'd function MB_GetString, if available */
				pmb_getstring p = (pmb_getstring)GetProcAddress(hLibUser32, "MB_GetString");

				if (p)
				{
					LPCWSTR text = NULL;
					h = GetDlgItem(hwnd, IDC_BUTTON_HELP);
					text = p(8); /* IDHELP */
					if (text && h)
					{
						SetWindowTextW(h, text);
					}

					h = GetDlgItem(hwnd, IDC_BUTTON_OK);
					text = p(0); /* IDOK */
					if (text && h)
					{
						SetWindowTextW(h, text);
					}

					h = GetDlgItem(hwnd, IDC_BUTTON_CANCEL);
					text = p(1); /* IDCANCEL */
					if (text && h)
					{
						SetWindowTextW(h, text);
					}
				} /* if (p) */
			} /* if (hLibUser32) */

			if (GetDlgCtrlID((HWND)wParam) != IDC_INPUT)
			{
				SetFocus(GetDlgItem(hwnd, IDC_INPUT));
				/** 
				 * Return false only to override the default behavior of DialogBoxes in Windows:
				 * The system assigns the default keyboard focus only if the dialog box procedure
				 * returns TRUE, and we don't want that.
				 */
				return FALSE;
			}

			return TRUE;
		} /* WM_INITDIALOG */

		case WM_COMMAND:
		{
			if ((LOWORD(wParam) == IDC_BUTTON_OK) ||
				(LOWORD(wParam) == IDC_BUTTON_CANCEL) ||
				(LOWORD(wParam) == IDC_BUTTON_HELP))
			{
				inputBoxStorage_t * is = (inputBoxStorage_t*)GetWindowLongPtr(hwnd, DWLP_USER);
				if (!is)
				{
					/* If we can't get our pointer, cancel this dialog. */
					EndDialog(hwnd, IDC_BUTTON_CANCEL);
				}

				switch (LOWORD(wParam))
				{
					case IDC_BUTTON_OK:
					{
						HWND h = GetDlgItem(hwnd, IDC_INPUT);
						
						int textLen = GetWindowTextLengthW(h);
						is->lpwstrReturnValue = new wchar_t[textLen + 1];
						GetWindowTextW(h, is->lpwstrReturnValue, textLen + 1);

						EndDialog(hwnd, LOWORD(wParam));
						break;
					} /* IDC_BUTTON_OK */

					case IDC_BUTTON_CANCEL:
					{
						EndDialog(hwnd, LOWORD(wParam));
						break;
					} /* IDC_BUTTON_CANCEL*/

					case IDC_BUTTON_HELP:
					{
						/*HWND help = HtmlHelp(
							GetDesktopWindow(),
							"c:\\Help.chm::/Intro.htm>Mainwin",
							HH_DISPLAY_TOPIC,
							NULL);*/
						/* TODO: support this HtmlHelp outdated thing */
						break;
					} /* IDC_BUTTON_HELP */
				}
			} /* if (LOWORD(wParam) == IDC_BUTTON_OK || IDC_BUTTON_CANCEL || IDC_BUTTON_HELP) */
			break;
		} /* WM_COMMAND */

		default:
		{
			return FALSE;
		}
	}

	return TRUE;
} /* InputBoxDlgProc */

extern
HMODULE g_dllModuleHandle;

/**
 * @brief			VB's input box implementation.
 * @param			pvarPrompt			Prompt to be used in the dialog box. Can be omitted.
 * @param			pvarTitle			Title to be used in the dialog box. Can be omitted.
 * @param			pvarDefaultValue	Default value to be used in the dialog box. Can be omitted.
 * @param			pvarX				X position of the dialog box. Can be omitted.
 * @param			pvarY				X position of the dialog box. Can be omitted.
 * @param			pvarHelpFile		Help file path. Can be omitted.
 * @param			pvarHelpContext		Help context ID (only if pvarHelpFile is specified). Can be omitted.
 * @returns			A BSTR with the user input text.
 */
EXPORT BSTR  __stdcall rtcInputBox(
	VARIANTARG		* pvarPrompt,
	VARIANTARG		* pvarTitle,
	VARIANTARG		* pvarDefaultValue,
	VARIANTARG		* pvarX,
	VARIANTARG		* pvarY,
	VARIANTARG		* pvarHelpFile,
	VARIANTARG		* pvarHelpContext
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvarPrompt %.8x, pvarTitle %.8x, pvarDefaultValue %.8x, pvarX %.8x, pvarY %.8x, pvarHelpFile %.8x, pvarHelpContext %.8x",
		(unsigned int)pvarPrompt,
		(unsigned int)pvarTitle,
		(unsigned int)pvarDefaultValue,
		(unsigned int)pvarX,
		(unsigned int)pvarY,
		(unsigned int)pvarHelpFile,
		(unsigned int)pvarHelpContext
	);

	inputBoxStorage_t st = { 0 };

	/* Setup the X value, if specified */
	if (pvarX->vt == VT_I2)
	{
		st.bHasX = true;
		st.lX = pvarX->iVal;
	}

	/* Setup the Y value, if specified. TODO: Support more types. */
	if (pvarY->vt == VT_I2)
	{
		st.bHasY = true;
		st.lY = pvarY->iVal;
	}

	/* Setup the HelpFile, if specified. TODO: Support more types. */
	if (pvarHelpFile)
	{
		st.lpwstrHelpFile = __vbaStrErrVarCopy(pvarHelpFile);
	}

	/* Setup the HelpContext, if specified */
	if (pvarHelpContext)
	{
		st.lpwstrHelpContext = __vbaStrErrVarCopy(pvarHelpContext);
	}

	/* Setup the Prompt, if specified */
	if (pvarPrompt)
	{
		st.lpwstrPrompt = __vbaStrErrVarCopy(pvarPrompt);
	}

	/* Setup the DefaultValue, if specified */
	if (pvarDefaultValue)
	{
		st.lpwstrDefaultValue = __vbaStrErrVarCopy(pvarDefaultValue);
	}

	/* Setup the Title, if specified */
	if (pvarTitle)
	{
		st.lpwstrTitle = __vbaStrErrVarCopy(pvarTitle);
	}

	/* Call the DialogBoxParamW and wait for the return value */
	INT_PTR ret = DialogBoxParamW(
		g_dllModuleHandle,
		MAKEINTRESOURCEW(ID_DIALOG_INPUTBOX),
		0,
		InputBoxDlgProc,
		(LPARAM)&st
	);

	/* Free the strings */
	if (st.lpwstrHelpFile)
	{
		__vbaFreeStr((BSTR*)&st.lpwstrHelpFile);
	}
	if (st.lpwstrHelpContext)
	{
		__vbaFreeStr((BSTR*)&st.lpwstrHelpContext);
	}
	if (st.lpwstrPrompt)
	{
		__vbaFreeStr((BSTR*)&st.lpwstrPrompt);
	}
	if (st.lpwstrDefaultValue)
	{
		__vbaFreeStr((BSTR*)&st.lpwstrDefaultValue);
	}
	if (st.lpwstrTitle)
	{
		__vbaFreeStr((BSTR*)&st.lpwstrTitle);
	}

	DEBUG_WIDE(
		"DialogBoxParamW ret = %.8x, GetLastError = %.8x, st.returnvalue %.8x",
		(unsigned int)ret,
		(unsigned int)GetLastError(),
		(unsigned int)st.lpwstrReturnValue
	);

	/* Check if the user accepted the dialog box or not */
	switch (ret)
	{
		case IDC_BUTTON_OK:
		{
			/* If there isn't a return vale, leave */
			if (!st.lpwstrReturnValue)
			{
				break;
			}

			/* Create a copy of the wchar_t value to a BSTR */
			BSTR ret = SysAllocString(st.lpwstrReturnValue);			
			/* Delete the wchar_t array */
			delete[] st.lpwstrReturnValue;

			return ret;
		}
		case IDC_BUTTON_CANCEL:
		default:
		{
			break;
		}
	}

	/* Default return value for cancel, or error */
	return SysAllocString(L"");
} /* rtcInputBox */