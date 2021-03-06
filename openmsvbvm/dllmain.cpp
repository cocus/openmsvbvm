#include "vba_internal.h"

#include "vba_exception.h"

#include "vba_strManipulation.h"
#include "vba_objManipulation.h"
#include "vba_varManipulation.h"


#include "vba_structures.h"
#include "vba_CommandLine.h"

#include <windows.h>
#include <strsafe.h>

HINSTANCE g_dllModuleHandle;

/**
 * @brief			DLL Main entry point.
 * @param			hModule			HMODULE of the DLL.
 * @param			dwReason		Why this DllMain was called.
 * @param			lpReserved		Reserved for shady stuff.
 * @return			TRUE always.
 */
BOOL APIENTRY DllMain(
	HMODULE		hModule,
	DWORD		dwReason,
	LPVOID		lpReserved
)
{
	switch (dwReason)
	{
		case DLL_PROCESS_ATTACH:
		{
			CoInitialize(NULL);
			g_dllModuleHandle = hModule;
			break;
		}
		case DLL_THREAD_ATTACH:
		{
			break;
		}
		case DLL_THREAD_DETACH:
		{
			break;
		}
		case DLL_PROCESS_DETACH:
		{
			break;
		}
	}
	return TRUE;
} /* DllMain */

#define HIDWORD(d)			(d >> 16)

EXPORT __int64 __stdcall __allmul(
	__int64		i64A,
	__int64		i64B
)
{
	__int64 result;

	if (HIDWORD(i64A) | HIDWORD(i64B))
		result = i64A * i64B;
	else
		result = (unsigned int)i64B * (unsigned __int64)(unsigned int)i64A;
	return result;
}

EXPORT double __stdcall __vbaCyMul(
	double		dblA,
	double		dblB
)
{
	return dblA * dblB / 1000.0; /* Because VB treats the Currencies as 1000 times their value */
}

EXPORT unsigned char __fastcall __vbaUI1I2(
	unsigned __int16 uiIn
)
{
	if (uiIn > 0xFFu)
	{
		vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
	}

	return (unsigned char)uiIn;
}

EXPORT __declspec(naked) void _adj_fptan(void)
{
	_asm
	{
		fptan
		retn
	}
}

EXPORT __declspec(naked) void _adj_fpatan(void)
{
	_asm
	{
		fpatan
		retn
	}
}

EXPORT void __stdcall _CIatan(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIcos(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIexp(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIlog(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIsin(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIsqrt(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CItan(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m16i(uint16_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m64(uint64_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m32(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m32i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m16i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m32i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m32(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m64(uint64_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fprem(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fprem1(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_r(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall __vbaChkstk(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

/**
 * @brief			Returns the exact same value that was specified as an argument.
 * @param			lpPointer		Value (pointer) to return.
 * @return			lpPointer always.
 */
EXPORT void * __stdcall VarPtr(
	void * lpPointer
)
{
	return lpPointer;
}

/**
 * @brief			Dumps the VBHeader struct into the debug channel.
 * @param			pvbhVB			Pointer to a VBHeader struct.
 */
void dumpVBHeaderStruct(struct VBHeader * pvbhVB)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_DECLARE_ASCII_BUFFER_IF_NEEDED();

	DEBUG_WIDE("(%.8x)->szVbMagic = %c%c%c%c", (unsigned int)pvbhVB, pvbhVB->szVbMagic[0], pvbhVB->szVbMagic[1], pvbhVB->szVbMagic[2], pvbhVB->szVbMagic[3]);
	DEBUG_WIDE("(%.8x)->wRuntimeBuild = %.5d", (unsigned int)pvbhVB, (unsigned int)pvbhVB->wRuntimeBuild);
	DEBUG_ASCII("(%.8x)->szLangDll = %s", (unsigned int)pvbhVB, pvbhVB->szLangDll);
	DEBUG_ASCII("(%.8x)->szSecLangDll = %s", (unsigned int)pvbhVB, pvbhVB->szSecLangDll);
	DEBUG_WIDE("(%.8x)->wRuntimeRevision = %.5d", (unsigned int)pvbhVB, (unsigned int)pvbhVB->wRuntimeRevision);
	DEBUG_WIDE("(%.8x)->dwLCID = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->dwLCID);
	DEBUG_WIDE("(%.8x)->dwSecLCID = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->dwSecLCID);
	DEBUG_WIDE("(%.8x)->lpSubMain = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->lpSubMain);
	DEBUG_WIDE("(%.8x)->lpProjectData = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->lpProjectData);
	DEBUG_WIDE("(%.8x)->fMdlIntCtls = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->fMdlIntCtls);
	DEBUG_WIDE("(%.8x)->fMdlIntCtls2 = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->fMdlIntCtls2);
	DEBUG_WIDE("(%.8x)->dwThreadFlags = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->dwThreadFlags);
	DEBUG_WIDE("(%.8x)->dwThreadCount = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->dwThreadCount);
	DEBUG_WIDE("(%.8x)->wFormCount = %.5d", (unsigned int)pvbhVB, (unsigned int)pvbhVB->wFormCount);
	DEBUG_WIDE("(%.8x)->wExternalCount = %.5d", (unsigned int)pvbhVB, (unsigned int)pvbhVB->wExternalCount);
	DEBUG_WIDE("(%.8x)->dwThunkCount = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->dwThunkCount);
	DEBUG_WIDE("(%.8x)->lpGuiTable = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->lpGuiTable);
	DEBUG_WIDE("(%.8x)->lpExternalTable = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->lpExternalTable);
	DEBUG_WIDE("(%.8x)->lpComRegisterData = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->lpComRegisterData);
	DEBUG_WIDE("(%.8x)->bSZProjectDescription = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->bSZProjectDescription);
	DEBUG_WIDE("(%.8x)->bSZProjectExeName = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->bSZProjectExeName);
	DEBUG_WIDE("(%.8x)->bSZProjectHelpFile = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->bSZProjectHelpFile);
	DEBUG_WIDE("(%.8x)->bSZProjectName = %.8x", (unsigned int)pvbhVB, (unsigned int)pvbhVB->bSZProjectName);
}

/**
 * @brief			Runs a VB application based on the specified VBHeader.
 * @param			pvbhVB			Pointer to a VBHeader struct.
 */
EXPORT void __stdcall ThunRTMain(
	struct VBHeader * pvbhVB
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"pvbhVB %.8x",
		(unsigned int)pvbhVB
	);

	// Spit out some debug info
	dumpVBHeaderStruct(pvbhVB);
	
	// Store the CommandLine
	LPWSTR cmdLineW = GetCommandLineW();
	rtcSetCommandLine(cmdLineW);

	DEBUG_WIDE(
		"SubMain() start"
	);

	FARPROC lpMain = (FARPROC)pvbhVB->lpSubMain;
	if (!lpMain)
	{
		MessageBoxW(0, L"No lpSubMain in VBHeader, cannot continue (for now)", nullptr, MB_ICONERROR);
	}
	else
	{
		// Check if the address of lpMain is a valid one to run
		if (IsBadCodePtr(lpMain))
		{
			MessageBoxW(0, L"lpSubMain IsBadCodePointer!, cannot continue", nullptr, MB_ICONERROR);
		}
		else
		{
			lpMain();
		}
	}

	DEBUG_WIDE(
		"SubMain() end"
	);

	ExitProcess(0);
} /* ThunRTMain */