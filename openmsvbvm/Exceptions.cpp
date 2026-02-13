#include "vba_internal.h"
#include "Exceptions.hpp"
#include "StringManipulation.hpp"

#include <vector>
#include <mutex>

#include <map>

std::mutex g_tLastErrorCode_mutex;
vbaExceptions g_iLastErrorCode = VBA_EXCEPTION_NO_ERROR;

std::mutex g_tLastExcepInfo_mutex;
EXCEPINFO g_tLastExcepInfo = { 0 };

static std::map<vbaExceptions, const wchar_t*> ExceptionDescriptionMapping = {
	{ VBA_EXCEPTION_NO_ERROR, L"No error" },
	{ VBA_EXCEPTION_RETURN_WITHOUT_GOSUB, L"Return without GoSub" },
	{ VBA_EXCEPTION_INVALID_PROCEDURE_CALL, L"Invalid procedure call" },
	{ VBA_EXCEPTION_OVERFLOW, L"Overflow" },
	{ VBA_EXCEPTION_OUT_OF_MEMORY, L"Out of memory" },
	{ VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE, L"Subscript out of range" },
	{ VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED, L"This array is fixed or temporarily locked" },
	{ VBA_EXCEPTION_DIVISION_BY_ZERO, L"Division by zero" },
	{ VBA_EXCEPTION_TYPE_MISMATCH, L"Type mismatch" },
	{ VBA_EXCEPTION_OUT_OF_STRING_SPACE, L"Out of string space" },
	{ VBA_EXCEPTION_EXPRESSION_TOO_COMPLEX, L"Expression too complex" },
	{ VBA_EXCEPTION_CANT_PERFORM_REQUESTED_OPERATION, L"Can't perform requested operation" },
	{ VBA_EXCEPTION_USER_INTERRUPT_OCCURRED, L"User interrupt occurred" },
	{ VBA_EXCEPTION_RESUME_WITHOUT_ERROR, L"Resume without error" },
	{ VBA_EXCEPTION_OUT_OF_STACK_SPACE, L"Out of stack space" },
	{ VBA_EXCEPTION_SUB_FUNCTION_OR_PROPERTY_NOT_DEFINED, L"Sub, Function, or Property not defined" },
	{ VBA_EXCEPTION_TOO_MANY_DLL_APPLICATION_CLIENTS, L"Too many DLL application clients" },
	{ VBA_EXCEPTION_ERROR_IN_LOADING_DLL, L"Error in loading DLL" },
	{ VBA_EXCEPTION_BAD_DLL_CALLING_CONVENTION, L"Bad DLL calling convention" },
	{ VBA_EXCEPTION_INTERNAL_ERROR, L"Internal error" },
	{ VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER, L"Bad file name or number" },
	{ VBA_EXCEPTION_FILE_NOT_FOUND, L"File not found" },
	{ VBA_EXCEPTION_BAD_FILE_MODE, L"Bad file mode" },
	{ VBA_EXCEPTION_FILE_ALREADY_OPEN, L"File already open" },
	{ VBA_EXCEPTION_DEVICE_IO_ERROR, L"Device I/O error" },
	{ VBA_EXCEPTION_FILE_ALREADY_EXISTS, L"File already exists" },
	{ VBA_EXCEPTION_BAD_RECORD_LENGTH, L"Bad record length" },
	{ VBA_EXCEPTION_DISK_FULL, L"Disk full" },
	{ VBA_EXCEPTION_INPUT_PAST_END_OF_FILE, L"Input past end of file" },
	{ VBA_EXCEPTION_BAD_RECORD_NUMBER, L"Bad record number" },
	{ VBA_EXCEPTION_TOO_MANY_FILES, L"Too many files" },
	{ VBA_EXCEPTION_DEVICE_UNAVAILABLE, L"Device unavailable" },
	{ VBA_EXCEPTION_PERMISSION_DENIED, L"Permission denied" },
	{ VBA_EXCEPTION_DISK_NOT_READY, L"Disk not ready" },
	{ VBA_EXCEPTION_CANT_RENAME_WITH_DIFFERENT_DRIVE, L"Can't rename with different drive" },
	{ VBA_EXCEPTION_PATH_FILE_ACCESS_ERROR, L"Path/File access error" },
	{ VBA_EXCEPTION_PATH_NOT_FOUND, L"Path not found" },
	{ VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET, L"Object variable or With block variable not set" },
	{ VBA_EXCEPTION_FOR_LOOP_NOT_INITIALIZED, L"For loop not initialized" },
	{ VBA_EXCEPTION_INVALID_PATTERN_STRING, L"Invalid pattern string" },
	{ VBA_EXCEPTION_INVALID_USE_OF_NULL, L"Invalid use of Null" },
	{ VBA_EXCEPTION_CANT_CALL_FRIEND_PROCEDURE, L"Can't call Friend procedure on an object that is not an instance of the defining class" },
	{ VBA_EXCEPTION_PROPERTY_OR_METHOD_CALL_INCLUDING_REFERENCE_TO_PRIVATE_OBJECT, L"A property or method call cannot include a reference to a private object, either as an argument or as a return value" },
	{ VBA_EXCEPTION_SYSTEM_DLL_COULD_NOT_BE_LOADED, L"System DLL could not be loaded" },
	{ VBA_EXCEPTION_CANT_USE_CHARACTER_DEVICES_IN_SPECIFIED_FILE_NAMES, L"Can't use character device names in specified file names" },
	{ VBA_EXCEPTION_INVALID_FILE_FORMAT, L"Invalid file format" },
	{ VBA_EXCEPTION_CANT_CREATE_NECESSARY_TEMPORARY_FILE, L"Can’t create necessary temporary file" },
	{ VBA_EXCEPTION_INVALID_FORMAT_IN_RESOURCE_FILE, L"Invalid format in resource file" },
	{ VBA_EXCEPTION_DATA_VALUE_NAMED_NOT_FOUND, L"Data value named not found" },
	{ VBA_EXCEPTION_ILLEGAL_PARAMETER_CANT_WRITE_ARRAYS, L"Illegal parameter; can't write arrays" },
	{ VBA_EXCEPTION_COULD_NOT_ACCESS_SYSTEM_REGISTRY, L"Could not access system registry" },
	{ VBA_EXCEPTION_COMPONENT_NOT_CORRECTLY_REGISTRED, L"Component not correctly registered" },
	{ VBA_EXCEPTION_COMPONENT_NOT_FOUND, L"Component not found" },
	{ VBA_EXCEPTION_COMPONENT_DID_NOT_RUN_CORRECTLY, L"Component did not run correctly" },
	{ VBA_EXCEPTION_OBJECT_ALREADY_LOADED, L"Object already loaded" },
	{ VBA_EXCEPTION_CANT_LOAD_OR_UNLOAD_THIS_OBJECT, L"Can't load or unload this object" },
	{ VBA_EXCEPTION_CONTROL_SPECIFIED_NOT_FOUND, L"Control specified not found" },
	{ VBA_EXCEPTION_OBJECT_WAS_UNLOADED, L"Object was unloaded" },
	{ VBA_EXCEPTION_UNABLE_TO_UNLOAD_WITHIN_THIS_CONTEXT, L"Unable to unload within this context" },
	{ VBA_EXCEPTION_SPECIFIED_FILE_OUT_OF_DATE_PROGRAM_REQUIRES_LATER_VERSION, L"The specified file is out of date. This program requires a later version" },
	{ VBA_EXCEPTION_SPECIFIED_OBJECT_CANT_BE_USED_AS_OWNER_FORM_FOR_SHOW, L"The specified object can't be used as an owner form for Show" },
	{ VBA_EXCEPTION_INVALID_PROPERTY_VALUE, L"Invalid property value" },
	{ VBA_EXCEPTION_INVALID_PROPERTY_ARRAY_INDEX, L"Invalid property-array index" },
	{ VBA_EXCEPTION_PROPERTY_SET_CANT_BE_EXECUTED_AT_RUN_TIME, L"Property Set can't be executed at run time" },
	{ VBA_EXCEPTION_PROPERTY_SET_CANT_BE_USED_WITH_A_READ_ONLY_PROPERTY, L"Property Set can't be used with a read-only property" },
	{ VBA_EXCEPTION_NEED_PROPERTY_ARRAY_INDEX, L"Need property-array index" },
	{ VBA_EXCEPTION_PROPERTY_SET_NOT_PERMITTED, L"Property Set not permitted" },
	{ VBA_EXCEPTION_PROPERTY_GET_CANT_BE_EXECUTED_AT_RUN_TIME, L"Property Get can't be executed at run time" },
	{ VBA_EXCEPTION_PROPERTY_GET_CANT_BE_EXECUTED_ON_WRITE_ONLY_PROPERTY, L"Property Get can't be executed on write-only property" },
	{ VBA_EXCEPTION_FORM_ALREADY_DISPLAYED_CANT_SHOW_MODALLY, L"Form already displayed; can't show modally" },
	{ VBA_EXCEPTION_CODE_MUST_CLOSE_TOPMOST_MODAL_FORM_FIRST, L"Code must close topmost modal form first" },
	{ VBA_EXCEPTION_PERMISSION_TO_USE_OBJECT_DENIED, L"Permission to use object denied" },
	{ VBA_EXCEPTION_PROPERTY_NOT_FOUND, L"Property not found" },
	{ VBA_EXCEPTION_PROPERTY_OR_METHOD_NOT_FOUND, L"Property or method not found" },
	{ VBA_EXCEPTION_OBJECT_REQUIRED, L"Object required" },
	{ VBA_EXCEPTION_INVALID_OBJECT_USE, L"Invalid object use" },
	{ VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT, L"Component can't create object or return reference to this object" },
	{ VBA_EXCEPTION_CLASS_DOESNT_SUPPORT_AUTOMATION, L"Class doesn't support Automation" },
	{ VBA_EXCEPTION_FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION, L"File name or class name not found during Automation operation" },
	{ VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_PROPERTY_OR_METHOD, L"Object doesn't support this property or method" },
	{ VBA_EXCEPTION_AUTOMATION_ERROR, L"Automation error" },
	{ VBA_EXCEPTION_CONNECTION_TO_TYPE_LIBRARY_OR_OBJECT_LIBRARY_FOR_REMOTE_PROCESS_HAS_BEEN_LOST, L"Connection to type library or object library for remote process has been lost" },
	{ VBA_EXCEPTION_AUTOMATION_DOESNT_HAVE_A_DEFAULT_VALUE, L"Automation object doesn't have a default value" },
	{ VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_ACTION, L"Object doesn't support this action" },
	{ VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_NAMED_ARGUMENTS, L"Object doesn't support named arguments" },
	{ VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_CURRENT_LOCALE_SETTING, L"Object doesn't support current locale setting" },
	{ VBA_EXCEPTION_NAMED_ARGUMENT_NOT_FOUND, L"Named argument not found" },
	{ VBA_EXCEPTION_ARGUMENT_NOT_OPTIONAL_OR_INVALID_PROPERTY_ASSIGNMENT, L"Argument not optional or invalid property assignment" },
	{ VBA_EXCEPTION_WRONG_NUMBER_OF_ARGUMENTS_OR_INVALID_PROPERTY_ASSIGNMENT, L"Wrong number of arguments or invalid property assignment" },
	{ VBA_EXCEPTION_OBJECT_NOT_A_COLLECTION, L"Object not a collection" },
	{ VBA_EXCEPTION_INVALID_ORDINAL, L"Invalid ordinal" },
	{ VBA_EXCEPTION_SPECIFIED_NOT_FOUND, L"Specified not found" },
	{ VBA_EXCEPTION_CODE_RESOURCE_NOT_FOUND, L"Code resource not found" },
	{ VBA_EXCEPTION_CODE_RESOURCE_LOCK_ERROR, L"Code resource lock error" },
	{ VBA_EXCEPTION_THIS_KEY_IS_ALREADY_ASSOCIATED_WITH_AN_ELEMENT_OF_THIS_COLLECTION, L"This key is already associated with an element of this collection" },
	{ VBA_EXCEPTION_VARIABLE_USES_A_TYPE_NOT_SUPPORTED_IN_VISUAL_BASIC, L"Variable uses a type not supported in Visual Basic" },
	{ VBA_EXCEPTION_THIS_COMPONENT_DOESNT_SUPPORT_THE_SET_OF_EVENTS, L"This component doesn't support the set of events" },
	{ VBA_EXCEPTION_INVALID_CLIPBOARD_FORMAT, L"Invalid Clipboard format" },
	{ VBA_EXCEPTION_METHOD_OR_DATA_MEMBER_NOT_FOUND, L"Method or data member not found" },
	{ VBA_EXCEPTION_THE_REMOTE_SERVER_MACHINE_DOES_NOT_EXIST_OR_IS_UNAVAILABLE, L"The remote server machine does not exist or is unavailable" },
	{ VBA_EXCEPTION_CLASS_NOT_REGISTREED_ON_LOCAL_MACHINE, L"Class not registered on local machine" },
	{ VBA_EXCEPTION_CANT_CREATE_AUTOREDRAW_IMAGE, L"Can't create AutoRedraw image" },
	{ VBA_EXCEPTION_INVALID_PICTURE, L"Invalid picture" },
	{ VBA_EXCEPTION_PRINTER_ERROR, L"Printer error" },
	{ VBA_EXCEPTION_PRINTER_DRIVER_DOES_NOT_SUPPORT_SPECIFIED_PROPERTY, L"Printer driver does not support specified property" },
	{ VBA_EXCEPTION_PROBLEM_GETTING_PRINTER_INFORMATION_FROM_SYSTEM, L"Problem getting printer information from the system. Make sure the printer is set up correctly" },
	{ VBA_EXCEPTION_INVALID_PICTURE_TYPE, L"Invalid picture type" },
	{ VBA_EXCEPTION_CANT_PRINT_FORM_IMAGE_TO_THIS_TYPE_OF_PRINTER, L"Can't print form image to this type of printer" },
	{ VBA_EXCEPTION_CANT_EMPTY_CLIPBOARD, L"Can't empty Clipboard" },
	{ VBA_EXCEPTION_CANT_OPEN_CLIPBOARD, L"Can't open Clipboard" },
	{ VBA_EXCEPTION_CANT_SAVE_FILE_TO_TEMP_DIRECTORY, L"Can't save file to TEMP directory" },
	{ VBA_EXCEPTION_SEARCH_TEXT_NOT_fOUND, L"Search text not found" },
	{ VBA_EXCEPTION_REPLACEMENTS_TOO_LONG, L"Replacements too long" },
	{ VBA_EXCEPTION_OUT_OF_MEMORY2, L"Out of memory" },
	{ VBA_EXCEPTION_NO_OBJECT, L"No object" },
	{ VBA_EXCEPTION_CLASS_IS_NOT_SET, L"Class is not set" },
	{ VBA_EXCEPTION_UNABLE_TO_ACTIVATE_OBJECT, L"Unable to activate object" },
	{ VBA_EXCEPTION_UNABLE_TO_CREATE_EMBEDDED_OBJECT, L"Unable to create embedded object" },
	{ VBA_EXCEPTION_ERROR_SAVING_TO_FILE, L"Error saving to file" },
	{ VBA_EXCEPTION_ERROR_LOADING_FROM_FILE, L"Error loading from file" }
};

/**
 * @brief			Convers a HRESULT code to VB's internal exception codes.
 * @param			hr			HRESULT code to convert.
 * @returns			vbaExceptions code representing the HRESULT provided.
 */
CEXTERN vbaExceptions __stdcall vbaErrorFromHRESULT(
	HRESULT		hr
)
{
	// TODO: use a map
	switch (hr)
	{
		// STG error codes
		case STG_E_INVALIDFUNCTION:			return VBA_EXCEPTION_INVALID_PROCEDURE_CALL;										// Unable to perform requested operation
		case STG_E_FILENOTFOUND:			return VBA_EXCEPTION_FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION;	// %1 could not be found.
		case STG_E_PATHNOTFOUND:			return VBA_EXCEPTION_PATH_NOT_FOUND;												// The path %1 could not be found.
		case STG_E_TOOMANYOPENFILES:		return VBA_EXCEPTION_TOO_MANY_FILES;												// There are insufficient resources to open another file.
		case STG_E_ACCESSDENIED:			return VBA_EXCEPTION_PERMISSION_DENIED;												// Access denied.
		case STG_E_INSUFFICIENTMEMORY:		return VBA_EXCEPTION_OUT_OF_MEMORY;													// There is insufficient memory available to complete operation.
		case STG_E_DISKISWRITEPROTECTED:	return VBA_EXCEPTION_PERMISSION_DENIED;												// Disk is write-protected. TODO: Check if it's the correct return value
		case STG_E_SEEKERROR:				return VBA_EXCEPTION_DEVICE_IO_ERROR;												// An error occurred during a seek operation. TODO: Check if it's the correct return value
		case STG_E_WRITEFAULT:				return VBA_EXCEPTION_DEVICE_IO_ERROR;												// A disk error occurred during a write operation. TODO: Check if it's the correct return value
		case STG_E_READFAULT:				return VBA_EXCEPTION_DEVICE_IO_ERROR;												// A disk error occurred during a read operation. TODO: Check if it's the correct return value
		case STG_E_FILEALREADYEXISTS:		return VBA_EXCEPTION_FILE_ALREADY_EXISTS;											// %1 already exists.
		case STG_E_MEDIUMFULL:				return VBA_EXCEPTION_DISK_FULL;														// There is insufficient disk space to complete operation.
		case STG_E_UNKNOWN:					return VBA_EXCEPTION_INTERNAL_ERROR;												// An unexpected error occurred.

		case STG_S_CONVERTED: // The underlying file was converted to compound file format.
		case STG_S_BLOCK: // The storage operation should block until more data is available.
		case STG_S_RETRYNOW: // The storage operation should retry immediately.
		case STG_S_MONITORING: // The notified event sink will not influence the storage operation.
		case STG_S_MULTIPLEOPENS: // Multiple opens prevent consolidated (commit succeeded).
		case STG_S_CONSOLIDATIONFAILED: // Consolidation of the storage file failed(commit succeeded).
		case STG_S_CANNOTCONSOLIDATE: // Consolidation of the storage file is inappropriate (commit succeeded).
		case STG_E_INVALIDHANDLE: // Attempted an operation on an invalid object.
		case STG_E_INVALIDPOINTER: // Invalid pointer error.
		case STG_E_NOMOREFILES: // There are no more entries to return.
		case STG_E_SHAREVIOLATION: // A share violation has occurred.
		case STG_E_LOCKVIOLATION: // A lock violation has occurred.
		case STG_E_INVALIDPARAMETER: // Invalid parameter error.
		case STG_E_PROPSETMISMATCHED: // Illegal write of non-simple property to simple property set.
		case STG_E_ABNORMALAPIEXIT: // An application programming interface (API) call exited abnormally.
		case STG_E_INVALIDHEADER: // The file %1 is not a valid compound file.
		case STG_E_INVALIDNAME: // The name %1 is not valid.
		case STG_E_UNIMPLEMENTEDFUNCTION: // That function is not implemented.
		case STG_E_INVALIDFLAG: // Invalid flag error.
		case STG_E_INUSE: // Attempted to use an object that is busy.
		case STG_E_NOTCURRENT: // The storage has been changed since the last commit.
		case STG_E_REVERTED: // Attempted to use an object that has ceased to exist.
		case STG_E_CANTSAVE: // Cannot save.
		case STG_E_OLDFORMAT: // The compound file %1 was produced with an incompatible version of storage.
		case STG_E_OLDDLL: // The compound file %1 was produced with a newer version of storage.
		case STG_E_SHAREREQUIRED: // Share.exe or equivalent is required for operation.
		case STG_E_NOTFILEBASEDSTORAGE: // Illegal operation called on non-file based storage.
		case STG_E_EXTANTMARSHALLINGS: // Illegal operation called on object with extant marshalings.
		case STG_E_DOCFILECORRUPT: // The docfile has been corrupted.
		case STG_E_BADBASEADDRESS: // OLE32.DLL has been loaded at the wrong address.
		case STG_E_DOCFILETOOLARGE: // The compound file is too large for the current implementation.
		case STG_E_NOTSIMPLEFORMAT: // The compound file was not created with the STGM_SIMPLE flag.
		case STG_E_INCOMPLETE: // The file download was aborted abnormally. The file is incomplete.
		case STG_E_TERMINATED: // The file download has been terminated.
		case STG_E_STATUS_COPY_PROTECTION_FAILURE: // Generic Copy Protection Error.
		case STG_E_CSS_AUTHENTICATION_FAILURE: // Copy Protection Error—DVD CSS Authentication failed.
		case STG_E_CSS_KEY_NOT_PRESENT: // Copy Protection Error—The given sector does not have a valid CSS key.
		case STG_E_CSS_KEY_NOT_ESTABLISHED: // Copy Protection Error—DVD session key not established.
		case STG_E_CSS_SCRAMBLED_SECTOR: // Copy Protection Error—The read failed because the sector is encrypted.
		case STG_E_CSS_REGION_MISMATCH: // Copy Protection Error—The current DVD's region does not correspond to the region setting of the drive.
		case STG_E_RESETS_EXHAUSTED: // Copy Protection Error—The drive's region setting might be permanent or the number of user resets has been exhausted.
		{
			return VBA_EXCEPTION_AUTOMATION_ERROR; // TODO: Is this right?
		}

		// TODO: FINISH THIS LIST
		case E_OUTOFMEMORY:			return VBA_EXCEPTION_OUT_OF_MEMORY;
		case E_INVALIDARG:			return VBA_EXCEPTION_AUTOMATION_ERROR;
		case E_NOTIMPL:				return VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_ACTION;
		case DISP_E_BADVARTYPE:		return VBA_EXCEPTION_AUTOMATION_ERROR;
		case DISP_E_ARRAYISLOCKED:	return VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED;
		default:					return VBA_EXCEPTION_INTERNAL_ERROR; // TODO: Is this right?
	}
} /* vbaErrorFromHRESULT */


/**
 * @brief			Return a constant string for the given VB exception.
 * @param			exception			The VB exception.
 */
static const wchar_t* GetVBExceptionDescription(vbaExceptions exception)
{
	auto iter = ExceptionDescriptionMapping.find(exception);

	if (iter != ExceptionDescriptionMapping.end())
	{
		return iter->second;
	}

	return nullptr;
} /* GetVBExceptionDescription */

/**
 * @brief			Raises an exception if the GetLastError() return value is diferent than S_OK.
 */
void __stdcall RaiseExceptionIfLastErrorIsSet(void)
{
	DWORD			lastError;

	lastError = GetLastError();

	switch (lastError)
	{
		case ERROR_PATH_NOT_FOUND:
		{
			vbaRaiseException(VBA_EXCEPTION_FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION);
		}

		case ERROR_NO_VOLUME_LABEL:
		case ERROR_MOD_NOT_FOUND:
		case ERROR_PROC_NOT_FOUND:
		{
			vbaRaiseException(VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_ACTION); // TODO: Check
		}

		default:
		{
			vbaRaiseException(VBA_EXCEPTION_FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION);
		}
	}
} /* RaiseExceptionIfLastErrorIsSet */

static ULONG_PTR const VBExceptionArguments[VB_EXCEPTION_NUMBER_OF_ARGUMENTS] = { VB_EXCEPTION_ARGUMENTS };

/**
 * @brief			Raises an exception using the VB specified exception number.
 * @param			exceptionCode		VB Exception code to raise.
 * @remark			This will call the SEH, which might terminate the app.
 */
CEXTERN void  vbaRaiseException(
	vbaExceptions		exceptionCode,
	EXCEPINFO const		*pExcepInfo
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"exceptionCode = %.4d, raising SEH Exception to die()",
		exceptionCode
	);

	{
		/* This is a little bit weird, I'm leaving the lock scope because RaiseException doesn't return. */
	//	std::unique_lock<std::mutex> lck(g_tLastErrorCode_mutex);
		g_iLastErrorCode = exceptionCode;
	}
	{
		/* This is a little bit weird, I'm leaving the lock scope because RaiseException doesn't return. */
	//	std::unique_lock<std::mutex> lck(g_tLastErrorCode_mutex);
		
		if (pExcepInfo)
		{
			g_tLastExcepInfo = *pExcepInfo;
		}
		else
		{
			EXCEPINFO tExcepInfo = { 0 };
			tExcepInfo.scode = exceptionCode;
			tExcepInfo.bstrDescription = SysAllocString(GetVBExceptionDescription(exceptionCode));
			g_tLastExcepInfo = tExcepInfo;
		}
	}


	// For some reason, VB raises exceptions with a hardcoded Exception Code, and exactly two
	// arguments, both of them 0xDEADCAFE.
	RaiseException(
		VB_MAGIC_EXCEPTION_CODE,
		0,
		VB_EXCEPTION_NUMBER_OF_ARGUMENTS,
		VBExceptionArguments
	);
} /* vbaRaiseException */

/**
 * @brief			Raises the overflow exception.
 */
EXPORT void __stdcall __vbaErrorOverflow(void)
{
	vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
} /* __vbaErrorOverflow */

typedef struct
{
	DWORD			hThreadId;
	DWORD			FS0;
	int				iUnk1;
} vbaOnErrorRegistredHandler;

std::mutex g_tRegistredOnErrorHandlers_mutex;
std::vector<vbaOnErrorRegistredHandler> g_tRegistredOnErrorHandlers;

EXPORT int __stdcall __vbaExitProc(void)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	std::vector<vbaOnErrorRegistredHandler>::const_iterator cit;
	{
		std::unique_lock<std::mutex> lck(g_tRegistredOnErrorHandlers_mutex);

		for (cit = g_tRegistredOnErrorHandlers.cbegin(); cit != g_tRegistredOnErrorHandlers.cend(); ++cit)
		{
			if (cit->hThreadId == ::GetCurrentThreadId())
			{
				DEBUG_WIDE(
					"Found current thread id in the registred on error handlers vector, removing it!"
				);

				g_tRegistredOnErrorHandlers.erase(cit);
				return 0;
			}
		}
	}

	return 0;
} /* __vbaExitProc */

EXPORT int __stdcall __vbaOnError(int iUnk1)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"iUnk1 %.8x",
		(int)iUnk1
	);

	std::vector<vbaOnErrorRegistredHandler>::const_iterator cit;
	{
		std::unique_lock<std::mutex> lck(g_tRegistredOnErrorHandlers_mutex);

		for (cit = g_tRegistredOnErrorHandlers.cbegin(); cit != g_tRegistredOnErrorHandlers.cend(); ++cit)
		{
			if (cit->hThreadId == ::GetCurrentThreadId())
			{
				DEBUG_WIDE(
					"Found current thread id in the registred on error handlers vector!"
				);
				return 0;
			}
		}
	
		vbaOnErrorRegistredHandler tNewHandler;

		tNewHandler.hThreadId = ::GetCurrentThreadId();
		tNewHandler.iUnk1 = iUnk1;

		g_tRegistredOnErrorHandlers.push_back(tNewHandler);
	}
	return 0;
} /* __vbaOnError */

typedef struct
{
	DWORD						dwFlags1;
	DWORD						dwFlags2;
	void						* lpOnErrorReturnAddress;
	DWORD						dwFlags3;
} vbaProcedureOnErrorInfo;

typedef enum
{
	PROCEDURE_INFO_FLAG_ON_ERROR_SET = 0x30,
} VBA_PROCEDURE_INFO_FLAGS;

typedef struct
{
	WORD						wFlags;
	WORD						wFrameSize;
	WORD						wNull1;
	WORD						wNull2;
	DWORD						dwNull1;
	void						(*lpCleanupFunction)(void);
	vbaProcedureOnErrorInfo		* lpOnErrorInfo;
	LPVOID						* lpOnErrorNextInstructionPointer;
} vbaProcedureInfo;

typedef struct
{
	LPVOID						* uhm1;
	LPVOID						* uhm2;
	LPVOID						* ESP;
	vbaProcedureInfo			* lpProcedureInfo;
	DWORD						otherInfo1;
	DWORD						otherInfo2;
	DWORD						otherInfo3;
	DWORD						OnErrorNextInstructionIndex;
} vbaProcedureLocalStorage;

/**
 * @brief			Check if the passed exception record matches the VB pattern for exceptions.
 * @param			ExceptionRecord		The exception record from the SEH.
 * @return			true if this exception was caused by VB, false otherwhise.
 */
static bool IsVbaExceptionCode(
	struct _EXCEPTION_RECORD* ExceptionRecord
)
{
	if ((ExceptionRecord->ExceptionCode == VB_MAGIC_EXCEPTION_CODE)
		&& (ExceptionRecord->NumberParameters == VB_EXCEPTION_NUMBER_OF_ARGUMENTS)
		&& !memcmp(VBExceptionArguments, ExceptionRecord->ExceptionInformation, sizeof(VBExceptionArguments)))
	{
		return true;
	}

	return false;
} /* IsVbaExceptionCode */

// TODO: Make proper declaration of this
extern void GetVBProjectTitle(BSTR* rhs);

/**
 * @brief			Shows the "runtime error" message box when a VB exception is not handled.
 * @param			ExceptionNumber		VB Exception code.
 * @param			ExceptionRecord		The exception record from the SEH.
 * @param			EstablisherFrame	The establisher frame pointer from the SEH.
 * @param			ContextRecord		The context record structure from the SEH.
 * @remark			Application will exit after the user acknowledges the message box.
 */
static void ShowUnhandledVBException(
	struct _EXCEPTION_RECORD	*ExceptionRecord,
	void						*EstablisherFrame, 
	struct _CONTEXT				*ContextRecord)
{
	wchar_t buff[300] = { 0 };

	vbaExceptions exc = (vbaExceptions)g_tLastExcepInfo.scode;

	// Print the text as VB, but adds additional exception information.
	swprintf(
		buff,
		299,
		L"Run-time error '%d':\n\n%s\n\nEIP: %.8x\nEBP: %.8x\nESP: %.8x\nFrame: %.8p",
		(int)exc,
		g_tLastExcepInfo.bstrDescription ? g_tLastExcepInfo.bstrDescription : L"Unknown!",
		ContextRecord ? ContextRecord->Eip : 0,
		ContextRecord ? ContextRecord->Ebp : 0,
		ContextRecord ? ContextRecord->Esp : 0,
		EstablisherFrame);
	
	BSTR title = nullptr;
	GetVBProjectTitle(&title);

	MessageBoxW(NULL, buff, title, MB_ICONEXCLAMATION | MB_OK);

	__vbaFreeStr(&title);
} /* ShowUnhandledVBException */

/**
 * @brief			Calls a function with a specified EBP value.
 * @param			fn			Function to call.
 * @param			new_ebp		EBP value to use when calling the function.
 * @remark			This is used to call the "cleanup" function before resuming after an error, and it needs
 *					to be called with the original EBP value of the errored function, otherwise it will crash.
 */
#pragma warning( push )
#pragma warning( disable : 4731 ) // EBP is modified by this code, we know.
static void __stdcall CallFunctionWithEBP(void (*fn)(void), DWORD new_ebp)
{
	__asm {
		push    ebp				// Store the current EBP (which MSVC has decided its value)
		mov     eax, fn			// Grab the function pointer, from the passed arguments (stack)
		mov     ebp, new_ebp	// Set the EBP value for the function call, from the passed arguments (stack)
		call    eax				// Call the function with the provided EBP
		pop     ebp				// Restore this function's EBP
	}
}
#pragma warning( pop )

/**
 * @brief			SEH handler.
 * @param			All the parameters used in the SEH frame handler.
 * @remark			This will handle the SEH, and terminates the APP if not OnError is set (for now).
 */
#pragma runtime_checks( "", off ) // This code ends up modifying the ESP, we know.
EXPORT EXCEPTION_DISPOSITION __cdecl __vbaExceptHandler(
	struct _EXCEPTION_RECORD	*ExceptionRecord,
	void						*EstablisherFrame,
	struct _CONTEXT				*ContextRecord,
	void						*DispatcherContext
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ExceptionRecord %.8x, EstablisherFrame %.8x, ContextRecord %.8x, DispatcherContext %.8x",
		(unsigned int)ExceptionRecord,
		(unsigned int)EstablisherFrame,
		(unsigned int)ContextRecord,
		(unsigned int)DispatcherContext
	);

	DEBUG_WIDE(
		"ExceptionRecord->ExceptionFlags %.8x",
		(unsigned int)ExceptionRecord->ExceptionFlags
	);

	if (IsVbaExceptionCode(ExceptionRecord) == false)
	{
		DEBUG_WIDE(
			"Not a VB Exception, returning ExceptionContinueSearch"
		);
		return ExceptionContinueSearch;
	}

	if (!EstablisherFrame)
	{
		DEBUG_WIDE(
			"EstablisherFrame == NULL, returning ExceptionContinueSearch"
		);
		// This is not a VB exception if EstablisherFrame is null
		return ExceptionContinueSearch;
	}

	DEBUG_WIDE("VB Exception detected, g_iLastErrorCode = %.4x", g_iLastErrorCode);

	// Look in the registred OnErrorHandlers to check if an OnError was set
	std::vector<vbaOnErrorRegistredHandler>::const_iterator cit;
	{
		std::unique_lock<std::mutex> lck(g_tRegistredOnErrorHandlers_mutex);

		for (cit = g_tRegistredOnErrorHandlers.cbegin(); cit != g_tRegistredOnErrorHandlers.cend(); ++cit)
		{
			if (cit->hThreadId == ::GetCurrentThreadId())
			{
				DEBUG_WIDE(
					"Found current thread id in the registred on error handlers vector, removing it!"
				);

				g_tRegistredOnErrorHandlers.erase(cit);
				
				vbaProcedureLocalStorage* localStorage = (vbaProcedureLocalStorage*)EstablisherFrame;

				DEBUG_WIDE(
					"ESP %.8x, lpProcedureInfo %.8x",
					(unsigned int)localStorage->ESP,
					(unsigned int)localStorage->lpProcedureInfo
				);

				// Check if we have access to the ProcedureInfo struct and if it's on a valid memory location
				if (!localStorage->lpProcedureInfo)
				{
					DEBUG_WIDE(
						"lpProcedureInfo == NULL, returning ExceptionContinueSearch"
					);
					// Weird!
					return ExceptionContinueSearch;
				}

				if (IsBadReadPtr(localStorage->lpProcedureInfo, sizeof(vbaProcedureInfo)))
				{
					DEBUG_WIDE(
						"lpProcedureInfo IsBadReadPtr!, returning ExceptionContinueSearch"
					);
					return ExceptionContinueSearch;
				}

				// We can restore the original EBP by obtainging the base EBP on this frame, and the "frame size" stored
				// on the ProcedureInfo structure, generated by the compiler for each function.
				ContextRecord->Ebp = (DWORD)(&localStorage->lpProcedureInfo) + localStorage->lpProcedureInfo->wFrameSize;

				// Original ESP is saved on the localStorage structure.
				ContextRecord->Esp = (DWORD)(localStorage->ESP);

				// If there's a cleanup function, call it now, before returning to the VB code.
				if (localStorage->lpProcedureInfo->lpCleanupFunction)
				{
					DEBUG_WIDE(
						"Calling cleanup function before resuming, function at %.8p",
						localStorage->lpProcedureInfo->lpCleanupFunction
					);

					// It's important to use the original function's EBP (as this function uses it to get variable's addresses
					// as offsets from it). It's the same EBP as the function uses normally.
					if (IsBadCodePtr((FARPROC)localStorage->lpProcedureInfo->lpCleanupFunction))
					{
						DEBUG_WIDE(
							"localStorage->lpProcedureInfo->lpCleanupFunction IsBadCodePtr!, returning ExceptionContinueSearch"
						);
						return ExceptionContinueSearch; // Weird!
					}
					CallFunctionWithEBP(localStorage->lpProcedureInfo->lpCleanupFunction, ContextRecord->Ebp);
				}

				if ((localStorage->lpProcedureInfo->wFlags & PROCEDURE_INFO_FLAG_ON_ERROR_SET))
				{
					// If there's a pointer to an "OnErrorInfo" structure, it means this exception is caught inside
					// an "On Error GoTo xxxx".
					if (localStorage->lpProcedureInfo->lpOnErrorInfo)
					{
						// Use the on-error resume address
						ContextRecord->Eip = (DWORD)(localStorage->lpProcedureInfo->lpOnErrorInfo->lpOnErrorReturnAddress);

						DEBUG_WIDE(
							"On Error GoTo detected, setting EBP to %.8x, ESP to %.8x, EIP to %.8x and returning",
							ContextRecord->Ebp,
							ContextRecord->Esp,
							ContextRecord->Eip
						);
					}
					// Otherwise, this exception is caught under an "On Error Resume Next".
					else
					{
						// VB updates this index prior to calls to functions that could cause exceptions. Each potential
						// call to a function that could cause an exception is counted, starting from 1.
						DWORD err_level = localStorage->OnErrorNextInstructionIndex;

						// VB generates an array of all the pointers to the next instruction after calls to functions that
						// could potentially cause an exception. First element on this list is the count of how many pointers
						// are stored.
						ContextRecord->Eip = (DWORD)(localStorage->lpProcedureInfo->lpOnErrorNextInstructionPointer[err_level + 1]);

						DEBUG_WIDE(
							"On Error Resume Next detected, block index %d, setting EBP to %.8x, ESP to %.8x, EIP to %.8x and returning",
							err_level + 1,
							ContextRecord->Ebp,
							ContextRecord->Esp,
							ContextRecord->Eip
						);
					}

					// Check if the on-error resume address is able to run
					if (IsBadCodePtr((FARPROC)ContextRecord->Eip))
					{
						DEBUG_WIDE(
							"ContextRecord->Eip IsBadCodePtr!, returning ExceptionContinueSearch"
						);
						return ExceptionContinueSearch; // Weird!
					}

					DEBUG_WIDE(
						"ContextRecord->Eip IsBadCodePtr is not bad code, continuing"
					);
					return ExceptionContinueExecution;
				}
				else
				{
					DEBUG_WIDE("not sure about this case");
					return ExceptionContinueSearch; // Weird!
				}
			}
		}
	}

	// Nothing handled the exception, so catch it with the VB user error.
	DEBUG_WIDE("Unhandled VB exception!");
	ShowUnhandledVBException(ExceptionRecord, EstablisherFrame, ContextRecord);
	::ExitProcess(-1); // TODO: check the appropriate return code.

	return ExceptionContinueSearch;
} /* __vbaExceptHandler */
#pragma runtime_checks( "", restore )

//_chkstk - check stack upon procedure entry - from Visual Studio CRT source code
//
//Purpose:
//       Provide stack checking on procedure entry. Method is to simply probe
//       each page of memory required for the stack in descending order. This
//       causes the necessary pages of memory to be allocated via the guard
//       page scheme, if possible. In the event of failure, the OS raises the
//       _XCPT_UNABLE_TO_GROW_STACK exception.
//
//       NOTE:  Currently, the (EAX < _PAGESIZE_) code path falls through
//       to the "lastpage" label of the (EAX >= _PAGESIZE_) code path.  This
//       is small; a minor speed optimization would be to special case
//       this up top.  This would avoid the painful save/restore of
//       ecx and would shorten the code path by 4-6 instructions.
//
//Entry:
//       EAX = size of local frame
//
//Exit:
//       ESP = new stackframe, if successful
EXPORT __declspec(naked) void __vbaChkstk(void)
{
	__asm {

        push    ecx

; Calculate new TOS.

        lea     ecx, [esp] + 8 - 4      ; TOS before entering function + size for ret value
        sub     ecx, eax                ; new TOS

; Handle allocation size that results in wraparound.
; Wraparound will result in StackOverflow exception.

        sbb     eax, eax                ; 0 if CF==0, ~0 if CF==1
        not     eax                     ; ~0 if TOS did not wrapped around, 0 otherwise
        and     ecx, eax                ; set to 0 if wraparound

        mov     eax, esp                ; current TOS
        and     eax, not ( 0x1000 - 1) ; Round down to current page boundary

cs10:
        cmp     ecx, eax                ; Is new TOS
		bnd jb  short cs20              ; in probed page?
        mov     eax, ecx                ; yes.
        pop     ecx

		; VB differs here: it clears the new local frame!
		push	eax
		push	edi
		push	ecx
		lea     edi, [eax + 4]
		lea     ecx, [esp + 16]
		sub     ecx, edi				; ecx = re-computed size of local frame from the beginning of the function, no wraparounds
		shr     ecx, 2					; because rep stosd moves 4 bytes at a time, divide the size of the local frame by 4
		sub		ecx, 2					; don't 'nuke the return address and the pushed eax
		xor		eax, eax
		rep		stosd
		pop		ecx
		pop		edi
		pop		eax

		xchg    esp, eax				; update esp
		mov     eax, dword ptr[eax]		; get return address
		mov     dword ptr[esp], eax		; and put it at new TOS

		bnd ret

; Find next lower page and probe
cs20:
        sub     eax, 0x1000				; decrease by PAGESIZE
        test    dword ptr [eax],eax     ; probe page.
        jmp     short cs10
	}
} /* __vbaChkstk */

/**
 * @brief			FP exception handler.
 * @remark			Stores the FPU status and jumps into _FpException.
 */
#pragma fenv_access(on) // This function reads the floating point status.
EXPORT void __fastcall __vbaFPException(void)
{
	unsigned int fp_status;
	_statusfp2(&fp_status, NULL);

	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();
	DEBUG_WIDE(
		"statusfp = %d", fp_status
	);

	while(true)
	{
		/* Stack Error */
		if (fp_status & _SW_INVALID)
		{
			vbaRaiseException(VBA_EXCEPTION_EXPRESSION_TOO_COMPLEX);
		}
		/* Divide by zero */
		else if (fp_status & _SW_ZERODIVIDE)
		{
			vbaRaiseException(VBA_EXCEPTION_DIVISION_BY_ZERO);
		}
		/* Overflow */
		else if (fp_status & _SW_OVERFLOW)
		{
			vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
		}
		/* Default case */
		else
		{
			vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
		}
	}

} /* __vbaFPException */

/**
 * @brief			Maybe: Stores GetLastError() value into a local storage. TODO: Check what this function does.
 */
EXPORT void __stdcall __vbaSetSystemError(void)
{

} /* __vbaSetSystemError */
