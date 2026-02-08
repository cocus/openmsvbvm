
#include "vba_internal.h"
#include "Exceptions.hpp"

#include <vector>
#include <string>

/* this is horrible, figure out a better way to access the global vb header */
#include "vba_structures.h"
extern VBHeader* g_pvbhGlobal;

// MIDL-generated header for this object
#include "App.h"

extern  ULONG g_Components;		/* from DllObjectInterface.cpp */

// TODO: Make proper declaration of this
extern void GetVBProjectTitle(BSTR* rhs);

class _AppImpl : public _App
{
	LONG refCount = 1;

public:
	// IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(
		REFIID riid,
		void** ppvObject) override
	{
		if (!ppvObject) return E_POINTER;
		*ppvObject = nullptr;

		if (riid == IID_IUnknown || riid == IID__App)
		{
			*ppvObject = static_cast<_App*>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	ULONG STDMETHODCALLTYPE AddRef() override
	{
		InterlockedIncrement(&g_Components);
		return InterlockedIncrement(&refCount);
	}

	ULONG STDMETHODCALLTYPE Release() override
	{
		ULONG r = InterlockedDecrement(&refCount);
		InterlockedDecrement(&g_Components);
		if (r == 0)
		{
			delete this;
		}
		return r;
	}

	HRESULT __stdcall get_Path(BSTR* rhs)
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

	HRESULT __stdcall put_Path(BSTR rhs)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall get_EXEName(BSTR* rhs)
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

	HRESULT __stdcall put_EXEName(BSTR rhs)
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

		// Apparently, the dwHandle has to be zero per the docs:
		// > A pointer to a variable that the function sets to zero.
		if (dwSize == 0 || dwHandle != 0) return false;

		block.resize(dwSize);
		if (!GetFileVersionInfoW(wszFileName, 0, dwSize, block.data()))
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

	HRESULT __stdcall get_Title(BSTR* rhs)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);

		GetVBProjectTitle(rhs);

		return S_OK;
	}

	HRESULT __stdcall put_Title(BSTR rhs)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
		return E_NOTIMPL;
	}

	HRESULT __stdcall GetTypeInfoCount(UINT* pctInfo)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall GetTypeInfo(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgdispid)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		return E_NOTIMPL;
	}


	HRESULT _stdcall HctlDemandLoad(unsigned int* ctl)
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall ChkProp(unsigned int i, unsigned int* tagData)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall SetPropA(unsigned int i, unsigned int* tagData)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall GetPropA(unsigned int i, unsigned int* tagData)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall GetPropHsz(unsigned int i, unsigned char** hsz)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall LoadProp(unsigned int i, unsigned int* fref)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall SaveProp(unsigned int i, unsigned int* fref)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall GetPalette(void)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall Reset(void)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall get_DefaultProp(VARIANT* var)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall put_DefaultProp(VARIANT* var)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall get_000x(VARIANT* var)
	{
		return E_NOTIMPL;
	}
	HRESULT _stdcall put_000x(unsigned int i)
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall get_PrevInstance(VARIANT_BOOL* rhs)
	{
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing27(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_StartMode(
		/* [retval][out] */ short* rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing29(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_TaskVisible(
		/* [retval][out] */ VARIANT_BOOL* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_TaskVisible(
		/* [in] */ VARIANT_BOOL rhs) {
		return E_NOTIMPL;
	}


	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyTimeout(
		/* [retval][out] */ long* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyTimeout(
		/* [in] */ long rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyMsgTitle(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyMsgTitle(
		/* [in] */ BSTR rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyMsgText(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyMsgText(
		/* [in] */ BSTR rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleServerBusyRaiseError(
		/* [retval][out] */ VARIANT_BOOL* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleServerBusyRaiseError(
		/* [in] */ VARIANT_BOOL rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingTimeout(
		/* [retval][out] */ long* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingTimeout(
		/* [in] */ long rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingMsgTitle(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingMsgTitle(
		/* [in] */ BSTR rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_OleRequestPendingMsgText(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_OleRequestPendingMsgText(
		/* [in] */ BSTR rhs) {
		return E_NOTIMPL;
	}


	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Major(
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

	/* [helpstring] */ HRESULT __stdcall Missing47(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Minor(
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

	/* [helpstring] */ HRESULT __stdcall Missing49(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Revision(
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

	/* [helpstring] */ HRESULT __stdcall Missing51(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Comments(
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

	/* [helpstring] */ HRESULT __stdcall Missing53(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_CompanyName(
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

	/* [helpstring] */ HRESULT __stdcall Missing55(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_FileDescription(
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

	/* [helpstring] */ HRESULT __stdcall Missing57(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LegalCopyright(
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

	/* [helpstring] */ HRESULT __stdcall Missing59(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LegalTrademarks(
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

	/* [helpstring] */ HRESULT __stdcall Missing61(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_ProductName(
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

	/* [helpstring] */ HRESULT __stdcall Missing63(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_hInstance(
		/* [retval][out] */ long* rhs)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
		*rhs = (long)GetModuleHandleW(NULL);

		return S_OK;
	}

	/* [helpstring] */ HRESULT __stdcall Missing65(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_NonModalAllowed(
		/* [retval][out] */ VARIANT_BOOL* rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing67(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LogPath(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing69(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_LogMode(
		/* [retval][out] */ long* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_UnattendedApp(
		/* [retval][out] */ VARIANT_BOOL* rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing71(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_ThreadID(
		/* [retval][out] */ long* rhs)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ("rhs %.8x", (unsigned int)rhs);
		*rhs = (long)GetCurrentThreadId();

		return S_OK;
	}

	/* [helpstring] */ HRESULT __stdcall Missing73(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_HelpFile(
		/* [retval][out] */ BSTR* rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing75(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_HelpFile(
		/* [in] */ BSTR rhs) {
		return E_NOTIMPL;
	}

	/* [helpstring] */ HRESULT __stdcall Missing77(void)
	{
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_RetainedProject(
		/* [retval][out] */ VARIANT_BOOL* rhs) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE StartLogging(
		/* [in] */ BSTR LogTarget,
		/* [in] */ long LogModes) {
		return E_NOTIMPL;
	}

	/* [helpcontext][helpstring] */ HRESULT STDMETHODCALLTYPE LogEvent(
		/* [in] */ BSTR LogBuffer,
		/* [in] */ VARIANT EventType) {
		return E_NOTIMPL;
	}
};

HRESULT Create_App_Instance(IUnknown** ppApp)
{
	if (!ppApp) return E_POINTER;
	*ppApp = new _AppImpl();
	if (!*ppApp) return E_OUTOFMEMORY;
	(*ppApp)->AddRef();
	return S_OK;
}