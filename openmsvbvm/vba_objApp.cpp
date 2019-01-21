#include "vba_objApp.h"

#include "vba_internal.h"
#include "vba_exception.h"

HRESULT __stdcall CApp::get_Path(BSTR * rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	wchar_t drive[_MAX_DRIVE];
	wchar_t dir[_MAX_DIR];
	wchar_t wszFileName[MAX_PATH];

	GetModuleFileNameW(NULL, wszFileName, MAX_PATH);

	errno_t err;

	err = _wsplitpath_s(
		wszFileName,
		drive, _MAX_DRIVE,
		dir, _MAX_DIR,
		nullptr, 0,
		nullptr, 0
	);

	swprintf(
		wszFileName,
		MAX_PATH,
		L"%ls%ls",
		drive,
		dir
	);

	*rhs = SysAllocString(wszFileName);
	return S_OK;
}

HRESULT __stdcall CApp::put_Path(BSTR rhs)
{
	return E_NOTIMPL;
}

HRESULT __stdcall CApp::get_EXEName(BSTR * rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	wchar_t filename[_MAX_FNAME];
	wchar_t wszFileName[MAX_PATH];

	GetModuleFileNameW(NULL, wszFileName, MAX_PATH);

	errno_t err;

	err = _wsplitpath_s(
		wszFileName,
		nullptr, 0,
		nullptr, 0,
		filename, _MAX_FNAME,
		nullptr, 0
	);

	*rhs = SysAllocString(filename);
	return S_OK;
}

HRESULT __stdcall CApp::put_EXEName(BSTR rhs)
{
	return E_NOTIMPL;
}

HRESULT __stdcall CApp::get_Title(BSTR * rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);



	/*DWORD dwHandle, dwLen;
	UINT BufLen;
	LPTSTR lpData, lpBuffer, LibName = argv[1];
	VS_FIXEDFILEINFO *pFileInfo;

	LibName = fullpath(argv[1]); // try first the current directory
	dwLen = GetFileVersionInfoSize(LibName, &dwHandle);
	if (!dwLen) {
		free(LibName);
		LibName = searchpath(argv[1], ".dll");
		dwLen = GetFileVersionInfoSize(LibName, &dwHandle);
	}
	if (!dwLen) {
		free(LibName);
		LibName = argv[1];
		dwLen = GetFileVersionInfoSize(LibName, &dwHandle);
	}
	printf("Library:              %s\n", LibName);
	if (!dwLen) {
		printf("VersionInfo           not found\n");
		return -1;
	}
	lpData = (LPTSTR)malloc(dwLen);
	if (!lpData) {
		perror("malloc");
		return -1;
	}
	if (!GetFileVersionInfo(LibName, dwHandle, dwLen, lpData)) {
		free(lpData);
		printf("VersionInfo:          not found\n");
		return -1;
	}
	if (!VerQueryValue(lpData, "\\", (LPVOID)&pFileInfo, (PUINT)&BufLen)) {
		printf("VersionInfo:          not found\n");
	}
	else {
		printf("MajorVersion:         %d\n", HIWORD(pFileInfo->dwFileVersionMS));
		printf("MinorVersion:         %d\n", LOWORD(pFileInfo->dwFileVersionMS));
		printf("BuildNumber:          %d\n", HIWORD(pFileInfo->dwFileVersionLS));
		printf("RevisionNumber (QFE): %d\n", LOWORD(pFileInfo->dwFileVersionLS));
	}

	if (!VerQueryValue(lpData, TEXT("\\StringFileInfo\\040904E4\\FileVersion"),
		(LPVOID)&lpBuffer, (PUINT)&BufLen)) {
		// language ID 040904E4: U.S. English, char set = Windows, Multilingual
		printf("FileVersion:          not found\n");
	}
	else
		printf("FileVersion:          %s\n", lpBuffer);

















	// Structure used to store enumerated languages and code pages.
	HRESULT hr;

	struct LANGANDCODEPAGE {
		WORD wLanguage;
		WORD wCodePage;
	} *lpTranslate;


	// Read the list of languages and code pages.


	VerQueryValue(pBlock,
		TEXT("\VarFileInfo\Translation"),
		(LPVOID*)&lpTranslate,
		&cbTranslate);


	// Read the file description for each language and code page.


	for (i = 0; i < (cbTranslate / sizeof(struct LANGANDCODEPAGE)); i++)
	{
		hr = StringCchPrintf(SubBlock, 50,
			TEXT("\StringFileInfo\%04x%04x\FileDescription"),
			lpTranslate[i].wLanguage,
			lpTranslate[i].wCodePage);
		if (FAILED(hr))
		{
			// TODO: write error handler.
		}


		// Retrieve file description for language and code page "i".
		VerQueryValue(pBlock,
			SubBlock,
			&lpBuffer,
			&dwBytes);
	}










	*/





	*rhs = SysAllocString(L"This should be the title!");
	return S_OK;
}

HRESULT __stdcall CApp::put_Title(BSTR rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
	return S_OK;
}


HRESULT __stdcall CApp::QueryInterface(
	REFIID riid,
	void **ppObj
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("riid , ppObj %.8x", (unsigned int)ppObj);
	if (riid == IID_IUnknown)
	{
		*ppObj = static_cast<void*>(this);
		AddRef();
		return S_OK;
	}
	if (riid == IID_IApp)
	{
		*ppObj = static_cast<void*>(this);
		AddRef();
		return S_OK;
	}
	//
	// if control reaches here then , let the client know that
	// we do not satisfy the required interface
	//
	*ppObj = NULL;
	return E_NOINTERFACE;
}

ULONG __stdcall CApp::AddRef()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	return InterlockedIncrement(&m_nRefCount);
}

ULONG __stdcall CApp::Release()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	long nRefCount = 0;
	nRefCount = InterlockedDecrement(&m_nRefCount);
	if (nRefCount == 0) delete this;
	return nRefCount;
}

HRESULT __stdcall CApp::GetTypeInfoCount(UINT * pctInfo)
{
	return E_NOTIMPL;
}

HRESULT __stdcall CApp::GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo ** pptinfo)
{
	return E_NOTIMPL;
}

HRESULT __stdcall CApp::GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames, UINT cNames, LCID lcid, DISPID * rgdispid)
{
	return E_NOTIMPL;
}

HRESULT __stdcall CApp::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pdispparams, VARIANT * pvarResult, EXCEPINFO * pexcepinfo, UINT * puArgErr)
{
	return E_NOTIMPL;
}


HRESULT _stdcall CApp::HctlDemandLoad(unsigned int * ctl)
{
	return E_NOTIMPL;
}

HRESULT _stdcall CApp::ChkProp(unsigned int i, unsigned int * tagData)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::SetPropA(unsigned int i, unsigned int * tagData)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::GetPropA(unsigned int i, unsigned int * tagData)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::GetPropHsz(unsigned int i, unsigned char ** hsz)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::LoadProp(unsigned int i, unsigned int * fref)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::SaveProp(unsigned int i, unsigned int * fref)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::GetPalette(void)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::Reset(void)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::get_DefaultProp(VARIANT * var)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::put_DefaultProp(VARIANT * var)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::get_000x(VARIANT * var)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CApp::put_000x(unsigned int i)
{
	return E_NOTIMPL;
}


extern  ULONG g_Components;		/* from DllObjectInterface.cpp */

CApp::CApp() : m_nRefCount(1)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	InterlockedIncrement((LONG*)&g_Components);
}

CApp::~CApp()
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("");

	InterlockedDecrement((LONG*)&g_Components);
}

