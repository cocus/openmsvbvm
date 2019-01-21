#include "vba_internal.h"

#include "vba_exception.h"

#include "vba_strManipulation.h"

#include <strsafe.h>




#define VB_REGISTRY_PATH_BASE			L"Software\\VB and VBA Program Settings\\"
#define VB_REGISTRY_PATH_BASE_LEN		37
#define VB_REGISTRY_DEFAULT_HIVE		HKEY_CURRENT_USER


/**
 * @brief			Saves a setting in the Windows registry.
 * @param			bstrAppName		Application name to be used for the registry sub key.
 * @param			bstrSection		Section name to be used for the registry sub key.
 * @param			bstrKey			Key name for the registry sub key.
 * @param			bstrSetting		Value for the registry sub key.
 */
EXPORT void __stdcall rtcSaveSetting(
	BSTR		bstrAppName,
	BSTR		bstrSection,
	BSTR		bstrKey,
	BSTR		bstrSetting
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"strAppName '%ls', strSection '%ls', strKey '%ls', strSetting '%ls'",
		bstrAppName,
		bstrSection,
		bstrKey,
		bstrSetting
	);

	wchar_t			wcSubKey[MAX_PATH];

	if (strSafeGetLength(bstrAppName) + strSafeGetLength(bstrSection) + VB_REGISTRY_PATH_BASE_LEN + 1 > MAX_PATH)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	swprintf(
		wcSubKey,
		MAX_PATH,
		VB_REGISTRY_PATH_BASE L"%ls\\%ls",
		bstrAppName,
		bstrSection
	);

	HKEY			hkSettingKey;
	LSTATUS			lResult;

	/* Create the settings key under wcSubkey in HKEY_CURRENT_USER */
	lResult = RegCreateKeyW(
		VB_REGISTRY_DEFAULT_HIVE,
		wcSubKey,
		&hkSettingKey
	);

	if (lResult != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	/* Save the user-defined settings in the previously created key */
	lResult = RegSetValueExW(
		hkSettingKey,
		bstrKey,
		NULL,
		REG_SZ,
		(const BYTE*)bstrSetting,
		strSafeGetLength(bstrSetting)
	);

	if (lResult != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	/* Close the previously created key */
	lResult = RegCloseKey(hkSettingKey);

	if (lResult != S_OK)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	return;
}


/* The following functions regarding the registry were taken from https://docs.microsoft.com/en-us/windows/desktop/sysinfo/deleting-a-key-with-subkeys */
//*************************************************************
//
//  RegDelnodeRecurse()
//
//  Purpose:    Deletes a registry key and all its subkeys / values.
//
//  Parameters: hKeyRoot    -   Root key
//              lpSubKey    -   SubKey to delete
//
//  Return:     ERROR_SUCCESS if successful.
//              LRESULT if an error occurs.
//
//*************************************************************

LRESULT RegDelnodeRecurse(HKEY hKeyRoot, LPWSTR lpSubKey)
{
	LPWSTR lpEnd;
	LONG lResult;
	DWORD dwSize;
	wchar_t szName[MAX_PATH];
	HKEY hKey;
	FILETIME ftWrite;

	// First, see if we can delete the key without having
	// to recurse.

	lResult = RegDeleteKeyW(hKeyRoot, lpSubKey);

	if (lResult == ERROR_SUCCESS)
		return lResult;

	lResult = RegOpenKeyExW(hKeyRoot, lpSubKey, 0, KEY_READ, &hKey);

	if (lResult != ERROR_SUCCESS)
	{
		return lResult;
	}

	// Check for an ending slash and add one if it is missing.

	lpEnd = lpSubKey + lstrlenW(lpSubKey);

	if (*(lpEnd - 1) != TEXT('\\'))
	{
		*lpEnd = TEXT('\\');
		lpEnd++;
		*lpEnd = TEXT('\0');
	}

	// Enumerate the keys

	dwSize = MAX_PATH;
	lResult = RegEnumKeyExW(hKey, 0, szName, &dwSize, NULL,
		NULL, NULL, &ftWrite);

	if (lResult == ERROR_SUCCESS)
	{
		do {

			*lpEnd = TEXT('\0');
			StringCchCatW(lpSubKey, MAX_PATH * 2, szName);

			if (!RegDelnodeRecurse(hKeyRoot, lpSubKey)) {
				break;
			}

			dwSize = MAX_PATH;

			lResult = RegEnumKeyExW(hKey, 0, szName, &dwSize, NULL,
				NULL, NULL, &ftWrite);

		} while (lResult == ERROR_SUCCESS);
	}

	lpEnd--;
	*lpEnd = TEXT('\0');

	RegCloseKey(hKey);

	// Try again to delete the key.

	lResult = RegDeleteKeyW(hKeyRoot, lpSubKey);

	return lResult;
}

//*************************************************************
//
//  RegDelnode()
//
//  Purpose:    Deletes a registry key and all its subkeys / values.
//
//  Parameters: hKeyRoot    -   Root key
//              lpSubKey    -   SubKey to delete
//
//  Return:     ERROR_SUCCESS if successful.
//              LRESULT if an error occurs.
//
//*************************************************************

LRESULT RegDelnode(HKEY hKeyRoot, LPWSTR lpSubKey)
{
	wchar_t szDelKey[MAX_PATH * 2];

	StringCchCopyW(szDelKey, MAX_PATH * 2, lpSubKey);
	return RegDelnodeRecurse(hKeyRoot, szDelKey);

}

/**
 * @brief			Deletes a setting (or multiple settings) in the Windows registry.
 * @param			bstrAppName		Application name to be used for the registry sub key.
 * @param			vargSection		Section name to be used for the registry sub key.
 * @param			vargKey			Key name for the registry sub key.
 * @remark			This will recursively delete multiple values if no key is specified.
 *					Also, if no section is specified, it will delete all the sub-sections within the appname.
 */
EXPORT void __stdcall rtcDeleteSetting(
	BSTR			bstrAppName,
	VARIANTARG		vargSection,
	VARIANTARG		vargKey
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	wchar_t			wcSubKey[MAX_PATH];
	LSTATUS			lResult;
	HKEY			hkKey;
	BSTR			bstrSection;
	BSTR			bstrKey;

	if (strSafeGetLength(bstrAppName) == 0)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return;
	}

	if (vargKey.vt == VT_ERROR)
	{
		if (vargSection.vt == VT_ERROR)
		{
			/* Section and key were not specified, so erase all the app settings */
			DEBUG_WIDE(
				"Erase all app settings '%ls'",
				bstrAppName
			);

			/* Prepare the Key path */
			if (strSafeGetLength(bstrAppName) + VB_REGISTRY_PATH_BASE_LEN > MAX_PATH)
			{
				vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
				return;
			}

			swprintf(
				wcSubKey,
				MAX_PATH,
				VB_REGISTRY_PATH_BASE L"%ls",
				bstrAppName
			);

			/* And delete that Key recursively */
			LRESULT lResult = RegDelnode(VB_REGISTRY_DEFAULT_HIVE, wcSubKey);
			if (lResult != ERROR_SUCCESS)
			{
				vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
				return;
			}
		}
		else
		{
			bstrSection = __vbaStrVarCopy(&vargSection);

			/* Key not specified, so erease the entire section */
			DEBUG_WIDE(
				"Erase entire section settings '%ls', '%ls'",
				bstrAppName,
				bstrSection
			);

			/* Prepare the Key path */
			if (strSafeGetLength(bstrAppName) + strSafeGetLength(bstrSection) + VB_REGISTRY_PATH_BASE_LEN + 1 > MAX_PATH)
			{
				vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
				return;
			}

			swprintf(
				wcSubKey,
				MAX_PATH,
				VB_REGISTRY_PATH_BASE L"%ls\\%ls",
				bstrAppName,
				bstrSection
			);

			/* And delete that Key recursively */
			LRESULT lResult = RegDelnode(VB_REGISTRY_DEFAULT_HIVE, wcSubKey);
			if (lResult != ERROR_SUCCESS)
			{
				vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
				return;
			}
		}
	}
	else
	{
		bstrSection = __vbaStrVarCopy(&vargSection);
		bstrKey = __vbaStrVarCopy(&vargKey);

		/* Erase a specific Key */
		DEBUG_WIDE(
			"Erase specific key setting '%ls', '%ls', '%ls'",
			bstrAppName,
			bstrSection,
			bstrKey
		);

		/* Prepare the Key path */
		if (strSafeGetLength(bstrAppName) + strSafeGetLength(bstrSection) + VB_REGISTRY_PATH_BASE_LEN + 1 > MAX_PATH)
		{
			vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
			return;
		}

		swprintf(
			wcSubKey,
			MAX_PATH,
			VB_REGISTRY_PATH_BASE L"%ls\\%ls",
			bstrAppName,
			bstrSection
		);

		/* Open that key */
		lResult = RegOpenKeyExW(
			VB_REGISTRY_DEFAULT_HIVE,
			wcSubKey,
			0,
			KEY_WRITE,
			&hkKey
		);

		if (lResult != ERROR_SUCCESS)
		{
			vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
			return;
		}

		/* Delete the specified value inside that key */
		lResult = RegDeleteValueW(
			hkKey,
			bstrKey
		);

		/* Cleanup */
		RegCloseKey(hkKey);

		if (lResult != ERROR_SUCCESS)
		{
			/* And notify if there were some error */
			vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		}
	}
}

/**
 * @brief			Gets a setting in the Windows registry. If it doesn't exists,
					it will use the provided default value.
 * @param			bstrAppName			Application name to be used for the registry sub key.
 * @param			bstrSection			Section name to be used for the registry sub key.
 * @param			bstrKey				Key name for the registry sub key.
 * @param			vargDefaultValue	Default value used if the registry key doesn't exists.
 * @return			BSTR with the setting value, or the default converted to BSTR.
 */
EXPORT BSTR __stdcall rtcGetSetting(
	BSTR			bstrAppName,
	BSTR			bstrSection,
	BSTR			bstrKey,
	VARIANTARG		vargDefaultValue
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	wchar_t			wcSubKey[MAX_PATH];
	LSTATUS			lResult;
	HKEY			hkKey;

	/* Erase a specific Key */
	DEBUG_WIDE(
		"Erase specific key setting '%ls', '%ls', '%ls'",
		bstrAppName,
		bstrSection,
		bstrKey
	);

	/* Prepare the Key path */
	if (strSafeGetLength(bstrAppName) + strSafeGetLength(bstrSection) + VB_REGISTRY_PATH_BASE_LEN + 1 > MAX_PATH)
	{
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return NULL;
	}

	swprintf(
		wcSubKey,
		MAX_PATH,
		VB_REGISTRY_PATH_BASE L"%ls\\%ls",
		bstrAppName,
		bstrSection
	);

	/* Open that key */
	lResult = RegOpenKeyExW(
		VB_REGISTRY_DEFAULT_HIVE,
		wcSubKey,
		0,
		KEY_READ,
		&hkKey
	);
	BSTR			bstrRet;

	if (lResult != ERROR_SUCCESS)
	{
		bstrRet = __vbaStrVarCopy(&vargDefaultValue);
		if (!bstrRet)
		{
			vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		}
		return bstrRet;
	}

	DWORD dwType;
	DWORD dwDataSize;

	/* Get the specified value */
	lResult = RegQueryValueExW(
		hkKey,
		bstrKey,
		NULL,
		&dwType,
		NULL,
		&dwDataSize
	);

	if (lResult != ERROR_SUCCESS)
	{
		RegCloseKey(hkKey);

		bstrRet = __vbaStrVarCopy(&vargDefaultValue);
		if (!bstrRet)
		{
			vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		}
		return bstrRet;
	}

	if (dwType != REG_SZ)
	{
		RegCloseKey(hkKey);
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
		return NULL;
	}

	bstrRet = SysAllocStringLen(
		NULL,
		(dwDataSize / 2) - 1
	);

	if (bstrRet == NULL)
	{
		SysFreeString(bstrRet);
		RegCloseKey(hkKey);
		vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
		return NULL;
	}

	/* Get the specified value */
	lResult = RegQueryValueExW(
		hkKey,
		bstrKey,
		NULL,
		&dwType,
		reinterpret_cast<LPBYTE>(bstrRet),
		&dwDataSize
	);

	/* Cleanup */
	RegCloseKey(hkKey);

	if (lResult != ERROR_SUCCESS)
	{
		/* And notify if there were some error */
		vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
	}

	return bstrRet;
}