#include "vba_internal.h"
#include "Exceptions.hpp"

#include "DllObjectInterface.hpp"

ULONG g_ServerLocks = 0;
ULONG g_Components = 0;

GUID CLSID_PropertyBag = { 0xd5de8d20, 0x5bb8, 0x11d1, { 0xa1, 0xe3, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x31 } };


EXTERN_C const CLSID CLSID_Global;



extern HRESULT Create_VBGlobal_Instance(IUnknown** ppApp);


/**
 * @brief			Instantiates a new COM object via it's CoClass and Interface, if specified.
 * @param			pvbNewData		Pointer to a VB defined structure with info about the object.
 * @param			ppv				Pointer to an IUnknown pointer, where the created object will be stored.
 * @returns			HRESULT of the operation.
 */
EXPORT HRESULT __stdcall __vbaNew2(
	vba_new_data_arg_t* pvbNewData,
	IUnknown** ppv
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	
	HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

	DEBUG_WIDE(
		"pvbNewData %.8x, reval %.8x",
		(unsigned int)pvbNewData,
		(unsigned int)ppv
	);


	if (!pvbNewData)
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}

	if (pvbNewData->lpguidCoClass)
	{
		DEBUG_WIDE_GUID(*pvbNewData->lpguidCoClass, "CoClass");
	}

	if (pvbNewData->lpguidInterface)
	{
		DEBUG_WIDE_GUID(*pvbNewData->lpguidInterface, "Interface");
	}

	if (!pvbNewData->lpguidCoClass || !pvbNewData->lpguidInterface)
	{
		vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
		return hr;
	}


	if (pvbNewData->flags & 2)
	{
		/* Check if the required CLSID is one of the internal VB CLSIDs */
		if (*pvbNewData->lpguidCoClass == CLSID_Global)
		{
			DEBUG_WIDE(
				"Got VBGlobal CoClass, using internal class"
			);

			/* Forward the instantiation to the VBGlobal file */
			hr = Create_VBGlobal_Instance(ppv);
		}
		else
		{
			/* Fallback to the original VB invocation of win32 api */
			IUnknown* ppunk = nullptr;
			hr = GetActiveObject(*pvbNewData->lpguidCoClass, NULL, &ppunk);
			if (SUCCEEDED(hr))
			{
				if (ppunk != nullptr)
				{
					hr = ppunk->QueryInterface(*pvbNewData->lpguidInterface, (void**)ppv);
				}
				else
				{
					hr = E_OUTOFMEMORY;
				}
			}
		}

		if (!SUCCEEDED(hr))
		{
			vbaRaiseException(VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT);
			return hr;
		}
	}

	return hr;
}

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


_Check_return_
STDAPI  DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	if (!ppv) return E_POINTER;
	*ppv = nullptr;

	if (IsEqualCLSID(rclsid, CLSID_PropertyBag))
	{
		DEBUG_WIDE(
			"Requested PropertyBag class object, but not implemented"
		);
		// TODO: Implement PropertyBag
		return E_NOTIMPL;
	}
	else
	{
		DEBUG_WIDE_GUID(rclsid, "Requested CLSID not available");
	}

	return CLASS_E_CLASSNOTAVAILABLE;
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