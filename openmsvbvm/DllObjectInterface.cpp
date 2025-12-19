#include "vba_internal.h"
#include "vba_exception.h"

#include "DllObjectInterface.h"
#include "ClassFactory.h"

#include "vba_objVBGlobal.h"
#include "vba_objApp.h"

ULONG g_ServerLocks = 0;
ULONG g_Components = 0;

GUID CLSID_VBGlobal =	{ 0xfcfb3d23, 0xa0fa, 0x1068, { 0xa7, 0x38, 0x08, 0x00, 0x2b, 0x33, 0x71, 0xb5 } };
// 33ad4f78-6699-11cf-b70c-00aa0060d393
GUID CLSID_App =		{ 0xfcfb3d23, 0xa0fa, 0x1068, { 0xa7, 0x38, 0x08, 0x00, 0x2b, 0x33, 0x71, 0xb5 } }; // TODO: Fix this GUID


STDAPI  DllCanUnloadNow(void)
{
	HRESULT rc = E_UNEXPECTED;

	if ((g_ServerLocks == 0) && (g_Components == 0))
	{
		rc = S_OK;
	}
	else
	{
		rc = S_FALSE;
	}
	return rc;
};

CMasterFactory factory;

STDAPI DllGetClassObject(
	_In_ REFCLSID rclsid,
	_In_ REFIID riid,
	_Outptr_ LPVOID FAR* ppv
)
{
	HRESULT rc = CLASS_E_CLASSNOTAVAILABLE;

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (rclsid == CLSID_VBGlobal)
	{
		DEBUG_WIDE(
			"clsid == CLSID_VBGlobal"
		);
		rc = factory.CreateInstance(nullptr, riid, ppv);
	}
	else if (rclsid == CLSID_App)
	{
		DEBUG_WIDE(
			"clsid == CLSID_App"
		);
		rc = factory.CreateInstance(nullptr, riid, ppv);
	}

	return rc;
};

STDAPI VBDllGetClassObject(
	_In_ HMODULE* phModule,
	int lReserved,
	_In_ struct VBHeader* pVBHeader, 
	_In_ REFCLSID rclsid,
	_In_ REFIID riid,
	_Outptr_ LPVOID FAR* ppv
)
{
	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI VBDllCanUnloadNow(
	struct VBHeader *pNewVBHeader
)
{
	return S_FALSE;
}