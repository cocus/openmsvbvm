#include "vba_objApp.h"

#include "vba_internal.h"
#include "vba_exception.h"

#include <vector>
#include <string>

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

/* Helper functions to read version resource from current module */
static bool LoadModuleVersionBlock(std::vector<BYTE>& block)
{
	wchar_t wszFileName[MAX_PATH];
	GetModuleFileNameW(NULL, wszFileName, MAX_PATH);

	DWORD dwHandle = 0;
	DWORD dwSize = GetFileVersionInfoSizeW(wszFileName, &dwHandle);
	if (dwSize == 0) return false;

	block.resize(dwSize);
	if (!GetFileVersionInfoW(wszFileName, dwHandle, dwSize, block.data()))
	{
		block.clear();
		return false;
	}
	return true;
}

static bool QueryVersionString(const std::vector<BYTE>& block, const wchar_t* name, std::wstring& out)
{
	if (block.empty()) return false;

	// Try to read translation table
	struct LANGANDCODEPAGE { WORD wLanguage; WORD wCodePage; } *lpTranslate = nullptr;
	UINT cbTranslate = 0;
	if (VerQueryValueW(block.data(), L"\\VarFileInfo\\Translation", (LPVOID*)&lpTranslate, &cbTranslate) && cbTranslate >= sizeof(*lpTranslate))
	{
		// use first translation
		wchar_t subkey[64];
		swprintf(subkey, _countof(subkey), L"\\StringFileInfo\\%04x%04x\\%s",
			lpTranslate[0].wLanguage, lpTranslate[0].wCodePage, name);

		LPWSTR lpBuffer = nullptr;
		UINT cb = 0;
		if (VerQueryValueW(block.data(), subkey, (LPVOID*)&lpBuffer, &cb) && lpBuffer && cb)
		{
			out.assign(lpBuffer, cb);
			// trim possible trailing nulls
			if (!out.empty() && out.back() == L'\0') out.pop_back();
			return true;
		}
	}

	// Fallback to US English (040904E4)
	{
		wchar_t subkey[64];
		swprintf(subkey, _countof(subkey), L"\\StringFileInfo\\040904E4\\%s", name);
		LPWSTR lpBuffer = nullptr;
		UINT cb = 0;
		if (VerQueryValueW(block.data(), subkey, (LPVOID*)&lpBuffer, &cb) && lpBuffer && cb)
		{
			out.assign(lpBuffer, cb);
			if (!out.empty() && out.back() == L'\0') out.pop_back();
			return true;
		}
	}

	return false;
}

static bool QueryFixedFileInfo(const std::vector<BYTE>& block, VS_FIXEDFILEINFO*& pInfo)
{
	if (block.empty()) return false;
	UINT cb = 0;
	if (!VerQueryValueW(block.data(), L"\\", (LPVOID*)&pInfo, &cb) || pInfo == nullptr) return false;
	return true;
}


/* this is horrible, figure out a better way to access the global vb header */
#include "vba_structures.h"
extern VBHeader* g_pvbhGlobal;

HRESULT __stdcall CApp::get_Title(BSTR* rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	/* Yes, the "title" is stored in the "exe name" */
	const char* ascii = ((char*)g_pvbhGlobal + g_pvbhGlobal->bSZProjectExeName);
	const size_t sz = strlen(ascii);

	int wslen = MultiByteToWideChar(CP_ACP, 0, ascii, sz, 0, 0);
	BSTR bstr = SysAllocStringLen(0, wslen);
	MultiByteToWideChar(CP_ACP, 0, ascii, sz, bstr, wslen);

	*rhs = SysAllocString(bstr);

	SysFreeString(bstr);
	return S_OK;
}

HRESULT __stdcall CApp::put_Title(BSTR rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
	return E_NOTIMPL;
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

HRESULT _stdcall CApp::get_PrevInstance(VARIANT_BOOL* rhs)
{
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_StartMode(
	/* [retval][out] */ short* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_TaskVisible(
	/* [retval][out] */ VARIANT_BOOL* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_TaskVisible(
	/* [in] */ VARIANT_BOOL rhs) {
	return E_NOTIMPL;
}



/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleServerBusyTimeout(
	/* [retval][out] */ long* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleServerBusyTimeout(
	/* [in] */ long rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleServerBusyMsgTitle(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleServerBusyMsgTitle(
	/* [in] */ BSTR rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleServerBusyMsgText(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleServerBusyMsgText(
	/* [in] */ BSTR rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleServerBusyRaiseError(
	/* [retval][out] */ VARIANT_BOOL* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleServerBusyRaiseError(
	/* [in] */ VARIANT_BOOL rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleRequestPendingTimeout(
	/* [retval][out] */ long* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleRequestPendingTimeout(
	/* [in] */ long rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleRequestPendingMsgTitle(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleRequestPendingMsgTitle(
	/* [in] */ BSTR rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_OleRequestPendingMsgText(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_OleRequestPendingMsgText(
	/* [in] */ BSTR rhs) {
	return E_NOTIMPL;
}


/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_Major(
	/* [retval][out] */ short* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = 0;
		return S_OK;
	}

	VS_FIXEDFILEINFO* pInfo = nullptr;
	if (!QueryFixedFileInfo(block, pInfo))
	{
		*rhs = 0;
		return S_OK;
	}

	*rhs = (short)HIWORD(pInfo->dwFileVersionMS);
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_Minor(
	/* [retval][out] */ short* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = 0;
		return S_OK;
	}

	VS_FIXEDFILEINFO* pInfo = nullptr;
	if (!QueryFixedFileInfo(block, pInfo))
	{
		*rhs = 0;
		return S_OK;
	}

	*rhs = (short)LOWORD(pInfo->dwFileVersionMS);
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_Revision(
	/* [retval][out] */ short* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = 0;
		return S_OK;
	}

	VS_FIXEDFILEINFO* pInfo = nullptr;
	if (!QueryFixedFileInfo(block, pInfo))
	{
		*rhs = 0;
		return S_OK;
	}

	// Use LOWORD(dwFileVersionLS) as revision (QFE) by convention
	*rhs = (short)LOWORD(pInfo->dwFileVersionLS);
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_Comments(
	/* [retval][out] */ BSTR* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"Comments", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_CompanyName(
	/* [retval][out] */ BSTR* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"CompanyName", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_FileDescription(
	/* [retval][out] */ BSTR* rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"FileDescription", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_LegalCopyright(
	/* [retval][out] */ BSTR* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"LegalCopyright", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_LegalTrademarks(
	/* [retval][out] */ BSTR* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"LegalTrademarks", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_ProductName(
	/* [retval][out] */ BSTR* rhs) {
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

	std::vector<BYTE> block;
	if (!LoadModuleVersionBlock(block))
	{
		*rhs = SysAllocString(L"");
		return S_OK;
	}

	std::wstring value;
	QueryVersionString(block, L"ProductName", value);
	*rhs = SysAllocString(value.empty() ? L"" : value.c_str());
	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_hInstance(
	/* [retval][out] */ long* rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
	*rhs = (long)GetModuleHandleW(NULL);

	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_NonModalAllowed(
	/* [retval][out] */ VARIANT_BOOL* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_LogPath(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_LogMode(
	/* [retval][out] */ long* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_UnattendedApp(
	/* [retval][out] */ VARIANT_BOOL* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_ThreadID(
	/* [retval][out] */ long* rhs)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
	*rhs = (long)GetCurrentThreadId();

	return S_OK;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_HelpFile(
	/* [retval][out] */ BSTR* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE CApp::put_HelpFile(
	/* [in] */ BSTR rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE CApp::get_RetainedProject(
	/* [retval][out] */ VARIANT_BOOL* rhs) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE CApp::StartLogging(
	/* [in] */ BSTR LogTarget,
	/* [in] */ long LogModes) {
	return E_NOTIMPL;
}

/* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE CApp::LogEvent(
	/* [in] */ BSTR LogBuffer,
	/* [in] */ VARIANT EventType) {
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

