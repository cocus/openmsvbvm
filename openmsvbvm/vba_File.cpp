#include "vba_internal.h"
#include "vba_exception.h"

#include <cstring>
#include <string>
#include <map>

#include "vba_enums.h"

#include "vba_objManipulation.h"
#include "vba_varManipulation.h"

class vbaFileAbstraction
{
public:
	vbaFileAbstraction(
		std::wstring file,
		vbaFileOpenMode mode
	) : file(file), mode(mode)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ(
			"mode = %.8x",
			(unsigned int)mode
		);

		if (mode & VB_FMODE_ACCESS_WRITE)
		{
			_wfopen_s(&sysHandle, file.c_str(), L"wb+");
		}
		else //if (mode & VB_FMODE_ACCESS_READ) // TODO!!!
		{
			_wfopen_s(&sysHandle, file.c_str(), L"rb+");
		}
	}

	void get3(
		unsigned int	uiSize,
		char			*pData
	)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ(
			"uiSize %.8x, pData %.8x",
			(unsigned int)uiSize,
			(unsigned int)pData
		);

		/* This should not happen, but... */
		if (pData == nullptr)
		{
			return;
		}

		/* TODO: check: If the size is null, then we have a pointer to a string */
		if (uiSize == 0)
		{
			DEBUG_WIDE_OBJ(
				"uiSize == 0!"
			);

			return;
		}

		size_t ret = fread(pData, 1, uiSize, sysHandle);

		DEBUG_WIDE_OBJ(
			"fread wrote %.8x bytes, and we aimed for %.8x bytes",
			(unsigned int)ret,
			(unsigned int)uiSize
		);
	} /* get3 */

	void put3(
		unsigned int	uiSize,
		char			*pData
	)
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ(
			"uiSize %.8x, pData %.8x",
			(unsigned int)uiSize,
			(unsigned int)pData
		);

		/* This should not happen, but... */
		if (pData == nullptr)
		{
			return;
		}

		/* TODO: check: If the size is null, then we have a pointer to a string */
		if (uiSize == 0)
		{
			/* Get the true BSTR from the specified pointer */
			pData = (char*)(*(BSTR*)pData);

			if (pData == nullptr)
			{
				DEBUG_WIDE_OBJ(
					"pData = NULL, after de-referencing the original pointer"
				);
				return;
			}

			/* Get the size of the string */
			uiSize = SysStringLen((BSTR)pData);
			if (uiSize == 0)
			{
				DEBUG_WIDE_OBJ(
					"wcslen = 0, could not get the size of the buffer to write"
				);
				return;
			}
		}

		size_t ret = fwrite(pData, 1, uiSize, sysHandle);

		DEBUG_WIDE_OBJ(
			"fwrite wrote %.8x bytes, and we aimed for %.8x bytes",
			(unsigned int)ret,
			(unsigned int)uiSize
		);
	} /* put3 */

	unsigned long rtcFileLength()
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		unsigned long ulOriginalPos = ftell(sysHandle);

		fseek(sysHandle, 0, SEEK_END);

		unsigned long ulSize = ftell(sysHandle);

		fseek(sysHandle, ulOriginalPos, SEEK_SET);

		DEBUG_WIDE_OBJ(
			"ulOriginalPos %.8x, ulSize %.8x",
			ulOriginalPos,
			ulSize
		);

		return ulSize;

	} /* rtcFileLength */

	~vbaFileAbstraction()
	{
		DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

		DEBUG_WIDE_OBJ(
			"this->sysHandle %.8x, file '%ls'",
			(unsigned int)sysHandle,
			file.c_str()
		);

		if (sysHandle)
		{
			fclose(sysHandle);
			sysHandle = NULL;
		}
	}

private:
	FILE * sysHandle = nullptr;
	std::wstring file = L"";
	vbaFileOpenMode mode;
}; /* class vbaFileAbstraction */


/* TODO: This should be thread-dependant, and maybe add locks? */
static std::map<unsigned int, vbaFileAbstraction*> _vbaFileHandles;

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
static bool vbaFileGetObjectFromVBHandle(
	unsigned int			uiVBHandle,
	vbaFileAbstraction**	obj
)
{
	std::map<unsigned int, vbaFileAbstraction*>::iterator it;

	it = _vbaFileHandles.find(uiVBHandle);

	if (it != _vbaFileHandles.end())
	{
		*obj = it->second;

		return true;
	}

	return false;
} /* vbaFileGetObjectFromVBHandle */

/**
 * @brief			Tries to open a file, and assigns the VB Handle identifier to the local list.
 * @param			uiMode			Mode bitfield (see vbaFileOpenMode enum) specifying how to open the file.
 * @param			unknown			??? (seems to be always -1).
 * @param			uiVBHandle		VB file handle identifier for this file.
 * @param			bstrFileName	File path.
 * @returns			The length of the file path argument (minus one) on success.
 */
EXPORT unsigned int __stdcall __vbaFileOpen(
	unsigned int	uiMode,
	int				unknown,
	unsigned int	uiVBHandle,
	BSTR			bstrFileName
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"uiMode %.8x, unknown %.8x, uiVBHandle %.8x, file = '%ls'",
		uiMode,
		unknown,
		uiVBHandle,
		bstrFileName
	);

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
	{
		vbaRaiseException(VBA_EXCEPTION_FILE_ALREADY_OPEN);
	}
	else
	{
		/* TODO: create a factory for this! */
		obj = new vbaFileAbstraction(std::wstring(bstrFileName), (vbaFileOpenMode)uiMode);

		_vbaFileHandles.insert(std::pair<unsigned int, vbaFileAbstraction*>(
			uiVBHandle,
			obj
		));
	}

	return wcslen(bstrFileName) + 1;
} /* __vbaFileOpen */

/**
 * @brief			Closes a previously open VB file.
 * @param			uiVBHandle		VB file handle identifier for this file.
 */
EXPORT void __stdcall __vbaFileClose(
	int vbHandle
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"vbHandle = %.8x",
		(unsigned int)vbHandle
	);

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(vbHandle, &obj))
	{
		_vbaFileHandles.erase(vbHandle);
		delete obj;
	}
	else
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
	}
} /* __vbaFileClose */

/**
 * @brief			Gets the file size of a previously open VB file.
 * @param			uiVBHandle		VB file handle identifier for this file.
 * @returns			The file size on success, 0 otherwise.
 */
EXPORT unsigned long __stdcall rtcFileLength(
	unsigned int	uiVBHandle
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"uiVBHandle %.8x",
		(unsigned int)uiVBHandle
	);

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
	{
		return obj->rtcFileLength();
	}
	else
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
		return 0;
	}
} /* rtcFileLength */

/**
 * @brief			Sets the absoulte position of a previously open VB file.
 * @param			ulPos			New position of the file (absolute).
 * @param			uiVBHandle		VB file handle identifier for this file.
 * @returns			The previous file position on success, 0 otherwise.
 */
EXPORT unsigned long __stdcall __vbaFileSeek(
	unsigned long	ulPos,
	unsigned int	uiVBHandle
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"ulPos %.8x, uiVBHandle %.8x",
		ulPos,
		(unsigned int)uiVBHandle
	);

	if (ulPos < 1)
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_RECORD_NUMBER);
		return 0;
	}

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
	{
		//return obj->rtcFileLength();
	}
	else
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
	}
	return 0;

} /* __vbaFileSeek */

/**
 * @brief			Reads data from a previously open VB file.
 * @param			uiSize			Size of the data buffer, or zero for strings.
 * @param			*pData			Pointer to the destination data buffer.
 * @param			uiVBHandle		VB file handle identifier for this file.
 */
EXPORT void __stdcall __vbaGet3(
	unsigned int	uiSize,
	char			*pData,
	unsigned int	uiVBHandle
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"uiSize %.8x, pData %.8x, uiVBHandle %.8x",
		(unsigned int)uiSize,
		(unsigned int)pData,
		(unsigned int)uiVBHandle
	);

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
	{
		obj->get3(uiSize, pData);
	}
	else
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
	}
} /* __vbaGet3 */

/**
 * @brief			Writes data to a previously open VB file.
 * @param			uiSize			Size of the data buffer, or zero for strings.
 * @param			*pData			Pointer to the source data buffer.
 * @param			uiVBHandle		VB file handle identifier for this file.
 */
EXPORT void __stdcall __vbaPut3(
	unsigned int	uiSize,
	char			*pData,
	unsigned int	uiVBHandle
)
{
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		"uiSize %.8x, pData %.8x, uiVBHandle %.8x",
		(unsigned int)uiSize,
		(unsigned int)pData,
		(unsigned int)uiVBHandle
	);

	vbaFileAbstraction * obj;
	if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
	{
		obj->put3(uiSize, pData);
	}
	else
	{
		vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
	}
} /* __vbaPut3 */