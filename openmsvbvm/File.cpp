#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include <cstring>
#include <string>
#include <map>

#include "vba_enums.h"

#include "ObjectManipulation.hpp"
#include "VariantManipulation.hpp"

class vbaFileAbstraction
{
public:
    vbaFileAbstraction(std::wstring file, vbaFileOpenMode mode) : file(file), mode(mode)
    {

        LOG_OBJ(LOG_DEBUG, this) << L"mode = " << vbl::Hex((unsigned long)mode);

        errno_t err = EINVAL;
        if (mode & VB_FMODE_ACCESS_WRITE)
        {
            err = _wfopen_s(&sysHandle, file.c_str(), L"wb+");
        }
        else if (mode & VB_FMODE_OUTPUT)
        {
            err = _wfopen_s(&sysHandle, file.c_str(), L"wb");
        }
        else // if (mode & VB_FMODE_ACCESS_READ) // TODO!!!
        {
            err = _wfopen_s(&sysHandle, file.c_str(), L"rb+");
        }

        if (err != 0)
        {
            LOG_OBJ(LOG_DEBUG, this) << L"_wfopen_s failed, err = " << vbl::Hex((unsigned long)err);
            if (GetLastError() == ERROR_SHARING_VIOLATION)
            {
                vbaRaiseException(VBA_EXCEPTION_PERMISSION_DENIED);
            }
            else
            {
                vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
            }
            sysHandle = nullptr;
            return;
        }
    }

    void get3(unsigned int uiSize, char* pData)
    {

        LOG_OBJ(LOG_DEBUG, this) << L"uiSize " << vbl::Hex((unsigned long)uiSize) << L", pData " << vbl::Hex((unsigned long)pData);

        /* This should not happen, but... */
        if (pData == nullptr)
        {
            return;
        }

        /* TODO: check: If the size is null, then we have a pointer to a string */
        if (uiSize == 0)
        {
            LOG_OBJ(LOG_DEBUG, this) << L"uiSize == 0!";

            return;
        }

        size_t ret = fread(pData, 1, uiSize, sysHandle);

        LOG_OBJ(LOG_DEBUG, this) << L"fread wrote " << vbl::Hex((unsigned long)ret) << L" bytes, and we aimed for "
                                 << vbl::Hex((unsigned long)uiSize) << L" bytes";
    } /* get3 */

    void put3(unsigned int uiSize, char* pData)
    {

        LOG_OBJ(LOG_DEBUG, this) << L"uiSize " << vbl::Hex((unsigned long)uiSize) << L", pData " << vbl::Hex((unsigned long)pData);

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
                LOG_OBJ(LOG_DEBUG, this) << L"pData = NULL, after de-referencing the original pointer";
                return;
            }

            /* Get the size of the string */
            uiSize = SysStringLen((BSTR)pData);
            if (uiSize == 0)
            {
                LOG_OBJ(LOG_DEBUG, this) << L"wcslen = 0, could not get the size of the buffer to write";
                return;
            }
        }

        size_t ret = fwrite(pData, 1, uiSize, sysHandle);

        LOG_OBJ(LOG_DEBUG, this) << L"fwrite wrote " << vbl::Hex((unsigned long)ret) << L" bytes, and we aimed for "
                                 << vbl::Hex((unsigned long)uiSize) << L" bytes";
    } /* put3 */

    void print(const BSTR pData)
    {

        LOG_OBJ(LOG_DEBUG, this) << L"pData " << vbl::Hex((unsigned long)pData);

        /* This should not happen, but... */
        if (pData == nullptr)
        {
            return;
        }

        UINT uiSize = SysStringLen(pData);

        /* Get the size of the string */
        if (uiSize == 0)
        {
            LOG_OBJ(LOG_DEBUG, this) << L"wcslen = 0, could not get the size of the buffer to write";
            return;
        }

        /* Get how many bytes we'll need to allocate */
        int size_needed = WideCharToMultiByte(CP_ACP, 0, (LPCWCH)pData, uiSize, NULL, 0, 0, 0);

        char* buffer = new char[size_needed + 1];

        if (!buffer)
        {
            vbaRaiseException(VBA_EXCEPTION_OUT_OF_MEMORY);
            return;
        }

        int iChars = WideCharToMultiByte(CP_ACP, 0, (LPCWCH)pData, uiSize + 1, buffer, size_needed + 1, 0, 0);

        size_t ret = fwrite(buffer, 1, uiSize, sysHandle);
        fwrite("\r\n", 1, 2, sysHandle); // add the CR+LF

        delete[] buffer;

        LOG_OBJ(LOG_DEBUG, this) << L"fwrite wrote " << vbl::Hex((unsigned long)ret) << L" bytes, and we aimed for "
                                 << vbl::Hex((unsigned long)uiSize) << L" bytes";
    } /* put3 */

    unsigned long rtcFileLength()
    {

        unsigned long ulOriginalPos = ftell(sysHandle);

        fseek(sysHandle, 0, SEEK_END);

        unsigned long ulSize = ftell(sysHandle);

        fseek(sysHandle, ulOriginalPos, SEEK_SET);

        LOG_OBJ(LOG_DEBUG, this) << L"ulOriginalPos " << vbl::Hex((unsigned long)ulOriginalPos) << L", ulSize "
                                 << vbl::Hex((unsigned long)ulSize);

        return ulSize;

    } /* rtcFileLength */

    ~vbaFileAbstraction()
    {

        LOG_OBJ(LOG_DEBUG, this) << L"this->sysHandle " << vbl::Hex((unsigned long)sysHandle) << L", file '" << file.c_str() << L"'";

        if (sysHandle)
        {
            fclose(sysHandle);
            sysHandle = NULL;
        }
    }

private:
    FILE* sysHandle = nullptr;
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
static bool vbaFileGetObjectFromVBHandle(unsigned int uiVBHandle, vbaFileAbstraction** obj)
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
EXPORT unsigned int __stdcall __vbaFileOpen(unsigned int uiMode, int unknown, unsigned int uiVBHandle, BSTR bstrFileName)
{

    LOG(LOG_DEBUG) << L"uiMode " << vbl::Hex((unsigned long)uiMode) << L", unknown " << vbl::Hex((unsigned long)unknown)
                   << L", uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle) << L", file = '" << vbl::Bstr(bstrFileName) << L"'";

    vbaFileAbstraction* obj;
    if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
    {
        vbaRaiseException(VBA_EXCEPTION_FILE_ALREADY_OPEN);
    }
    else
    {
        /* TODO: create a factory for this! */
        obj = new vbaFileAbstraction(std::wstring(bstrFileName), (vbaFileOpenMode)uiMode);

        _vbaFileHandles.insert(std::pair<unsigned int, vbaFileAbstraction*>(uiVBHandle, obj));
    }

    return wcslen(bstrFileName) + 1;
} /* __vbaFileOpen */

/**
 * @brief			Closes a previously open VB file.
 * @param			uiVBHandle		VB file handle identifier for this file.
 */
EXPORT void __stdcall __vbaFileClose(int vbHandle)
{

    LOG(LOG_DEBUG) << L"vbHandle = " << vbl::Hex((unsigned long)vbHandle);

    vbaFileAbstraction* obj;
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
EXPORT unsigned long __stdcall rtcFileLength(unsigned int uiVBHandle)
{

    LOG(LOG_DEBUG) << L"uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle);

    vbaFileAbstraction* obj;
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
EXPORT unsigned long __stdcall __vbaFileSeek(unsigned long ulPos, unsigned int uiVBHandle)
{

    LOG(LOG_DEBUG) << L"ulPos " << vbl::Hex((unsigned long)ulPos) << L", uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle);

    if (ulPos < 1)
    {
        vbaRaiseException(VBA_EXCEPTION_BAD_RECORD_NUMBER);
        return 0;
    }

    vbaFileAbstraction* obj;
    if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
    {
        // return obj->rtcFileLength();
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
EXPORT void __stdcall __vbaGet3(unsigned int uiSize, char* pData, unsigned int uiVBHandle)
{

    LOG(LOG_DEBUG) << L"uiSize " << vbl::Hex((unsigned long)uiSize) << L", pData " << vbl::Hex((unsigned long)pData)
                   << L", uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle);

    vbaFileAbstraction* obj;
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
EXPORT void __stdcall __vbaPut3(unsigned int uiSize, char* pData, unsigned int uiVBHandle)
{

    LOG(LOG_DEBUG) << L"uiSize " << vbl::Hex((unsigned long)uiSize) << L", pData " << vbl::Hex((unsigned long)pData)
                   << L", uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle);

    vbaFileAbstraction* obj;
    if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
    {
        obj->put3(uiSize, pData);
    }
    else
    {
        vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
    }
} /* __vbaPut3 */

EXPORT int __vbaPrintFile(LPVOID descriptor, unsigned int uiVBHandle, const BSTR pData)
{

    LOG(LOG_DEBUG) << L"descriptor " << vbl::Hex((unsigned long)descriptor) << L", pData " << vbl::Hex((unsigned long)pData)
                   << L", uiVBHandle " << vbl::Hex((unsigned long)uiVBHandle);

    vbaFileAbstraction* obj;
    if (vbaFileGetObjectFromVBHandle(uiVBHandle, &obj))
    {
        obj->print(pData);
    }
    else
    {
        vbaRaiseException(VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER);
    }
    return 0;
}
