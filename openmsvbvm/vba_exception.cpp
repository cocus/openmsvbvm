#include "vba_internal.h"
#include "vba_exception.h"

/**
 * @brief			Convers a HRESULT code to VB's internal exception codes.
 * @param			hr			HRESULT code to convert.
 * @returns			vbaExceptions code representing the HRESULT provided.
 */
CEXTERN vbaExceptions __stdcall vbaErrorFromHRESULT(
	HRESULT		hr
)
{
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
 * @brief			Raise error if the GetLastError return value is diferent than S_OK.
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

/**
 * @brief			Raises an exception using the VB specified exception number.
 * @param			exceptionCode		VB Exception code to raise.
 * @remark			This will call the SEH, which might terminate the app.
 */
CEXTERN void  vbaRaiseException(
	vbaExceptions		exceptionCode
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"exceptionCode = %.4d, raising SEH Exception to die()",
		exceptionCode
	);

	RaiseException(
		exceptionCode,
		0,
		0,
		nullptr
	);
} /* vbaRaiseException */



bool onerrorset = false;

EXPORT int __stdcall __vbaExitProc(void)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"?"
	);

	onerrorset = false;

	return 0;
} /* __vbaExitProc */

EXPORT int __stdcall __vbaOnError(int iUnk1)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"iUnk1 %.8x",
		(int)iUnk1
	);

	onerrorset = true;

	return 0;
} /* __vbaOnError */

typedef struct
{
	uint32_t					dwFlags1;
	uint32_t					dwFlags2;
	void						* lpOnErrorReturnAddress;
	uint32_t					dwFlags3;
} vbaProcedureOnErrorInfo;

typedef enum
{
	PROCEDURE_INFO_FLAG_ON_ERROR_SET = 0x30,
} VBA_PROCEDURE_INFO_FLAGS;

typedef struct
{
	uint16_t					wFlags;
	uint16_t					wFrameSize;
	uint32_t					dwNull1;
	uint32_t					dwNull2;
	void						* lpSafeReturnAddr;
	vbaProcedureOnErrorInfo		* lpOnErrorInfo;
} vbaProcedureInfo;


typedef union
{
	vbaProcedureInfo			* pInfo;
	int							pEBPbase;
} vbaProcedureInfoUnion;

typedef struct
{
	void						* uhm1;
	void						* uhm2;
	void						* ESP;
	vbaProcedureInfoUnion		u;
	int							otherInfo1;
	int							otherInfo2;
} vbaProcedureLocalStorage;

/**
 * @brief			SEH handler.
 * @param			All the parameters used in the SEH frame handler.
 * @remark			This will handle the SEH, and terminates the APP if not OnError is set (for now).
 */
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

	// Temporary hack to handle OnError cases
	if (onerrorset)// && !(ExceptionRecord->ExceptionCode & EXCEPTION_NONCONTINUABLE))
	{
		vbaProcedureLocalStorage	* localStorage = (vbaProcedureLocalStorage*)EstablisherFrame;

		onerrorset = false;

		if (!localStorage)
		{
			DEBUG_WIDE(
				"localStorage == NULL, returning ExceptionContinueSearch"
			);
			return ExceptionContinueSearch;
		}

		DEBUG_WIDE(
			"Jumping to VB On-Error Handler, ESP %.8x, u.pInfo %.8x",
			(unsigned int)localStorage->ESP,
			(unsigned int)localStorage->u.pInfo
		);

		/* Check if we have access to the ProcedureInfo struct and if it's on a valid memory location */
		if (!localStorage->u.pInfo)
		{
			DEBUG_WIDE(
				"u.pInfo == NULL, returning ExceptionContinueSearch"
			);
			return ExceptionContinueSearch;
		}
		if (IsBadReadPtr(localStorage->u.pInfo, sizeof(vbaProcedureInfo)))
		{
			DEBUG_WIDE(
				"u.pInfo IsBadReadPtr!, returning ExceptionContinueSearch"
			);
			return ExceptionContinueSearch;
		}

		// This is a little bit of a magic thing, and I got most of this by just decompiling a dummy exe file
		// We can restore the original EBP by obtainging the base EBP on this frame, and it's size.
		ContextRecord->Ebp = (DWORD)(&localStorage->u.pEBPbase) + localStorage->u.pInfo->wFrameSize;
		// Original ESP is saved on the localStorage structure.
		ContextRecord->Esp = (DWORD)(localStorage->ESP);

		if (localStorage->u.pInfo->wFlags & PROCEDURE_INFO_FLAG_ON_ERROR_SET)
		{
			/* If we have a on-error resume address, use it */
			ContextRecord->Eip = (DWORD)(localStorage->u.pInfo->lpOnErrorInfo->lpOnErrorReturnAddress);

			DEBUG_WIDE(
				"On-Error resume address found, setting EBP to %.8x, ESP to %.8x, EIP to %.8x and returning",
				ContextRecord->Ebp,
				ContextRecord->Esp,
				ContextRecord->Eip
			);

			// Check if the on-error resume address is able to run
			if (IsBadCodePtr((FARPROC)ContextRecord->Eip))
			{
				DEBUG_WIDE(
					"ContextRecord->Eip IsBadCodePtr!, returning ExceptionContinueSearch"
				);
				return ExceptionContinueSearch;
			}

			return ExceptionContinueExecution;
		}
		else
		{
			/* There's no resume address, so call the cleanup function */
			ContextRecord->Eip = (DWORD)(localStorage->u.pInfo->lpSafeReturnAddr);

			DEBUG_WIDE(
				"No on-Error resume, setting EBP to %.8x, ESP to %.8x, EIP to %.8x and returning",
				ContextRecord->Ebp,
				ContextRecord->Esp,
				ContextRecord->Eip
			);

			// Check if the cleanup function address is able to run
			if (IsBadCodePtr((FARPROC)ContextRecord->Eip))
			{
				DEBUG_WIDE(
					"ContextRecord->Eip IsBadCodePtr!, returning ExceptionContinueSearch"
				);
				return ExceptionContinueSearch;
			}

			return ExceptionContinueExecution;
		}
	}

	return ExceptionContinueSearch;
} /* __vbaExceptHandler */

/**
 * @brief			Raises a VB error depending on the FPU exception type.
 * @param			fp_status		FPU status word.
 */
__declspec(naked) void __fastcall _FpExcecption(
	int			fp_status
)
{
	/* Stack Error */
	if (fp_status & 0x40)
	{
		vbaRaiseException(VBA_EXCEPTION_EXPRESSION_TOO_COMPLEX);
	}
	/* Divide by zero */
	else if (fp_status & 0x04)
	{
		vbaRaiseException(VBA_EXCEPTION_DIVISION_BY_ZERO);
	}
	/* Overflow */
	else if (fp_status & 0x08)
	{
		vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
	}
	/* Default case */
	else
	{
		vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
	}
} /* _FpExcecption */

/**
 * @brief			FP exception handler.
 * @remark			Stores the FPU status and jumps into _FpException.
 */
EXPORT __declspec(naked) void __fastcall __vbaFPException(void)
{
	__asm
	{
		fnstsw ax				/* Store FPU status word in AX register after checking for pending unmasked floating-point exceptions.*/
		jmp _FpExcecption
	}
} /* __vbaFPException */

/**
 * @brief			Maybe: Stores GetLastError() value into a local storage. TODO: Check what this function does.
 */
EXPORT void __stdcall __vbaSetSystemError(void)
{

} /* __vbaSetSystemError */
