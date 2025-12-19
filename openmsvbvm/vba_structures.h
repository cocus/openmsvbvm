#pragma once

//#include <windows.h>






struct MethodLinkNative
{
	char	jmpOpCode;
	int		jmpOffset;
};


struct OptionalObjectInfo
{
	DWORD dwObjectGuids;					/* How many GUIDs to Register. 2 = Designer */
	CLSID* clsidObjectClass;				/* Unique GUID of the Object */
	CLSID* clsidObjectInterface;			/* Unused */
	LPVOID lpuuidObjectTypes;				/* Pointer to Array of Object Interface GUIDs */
	DWORD dwObjectTypeGuids;				/* How many GUIDs in the Array above */
	LPVOID lpControls2;						/* Usually the same as lpControls */
	DWORD dwNull2;							/* Unused */
	LPVOID lpObjectGuid2;					/* Pointer to Array of Object GUIDs */
	DWORD dwControlCount;					/* Number of Controls in array below */
	LPVOID lpControls;						/* Pointer to Controls Array */
	WORD wEventCount;						/* Number of Events in Event Array */
	WORD wPCodeCount;						/* Number of P-Codes used by this Object */
	WORD bWInitializeEvent;					/* Offset to Initialize Event from Event Table */
	WORD bWTerminateEvent;					/* Offset to Terminate Event in Event Table */
	LPVOID lpEvents;						/* Pointer to Events Array */
	LPVOID lpBasicClassObject;				/* Pointer to in-memory Class Objects */
	DWORD dwNull3;							/* Unused */
	LPVOID lpIdeData;						/* Only valid in IDE */
	/*
		fDesigner As Long              ' 0x00 (0d) If this value is 2 then this object is a designer
	aObjectCLSID As Long           ' 0x04
	Null1 As Long                  ' 0x08
	aGuidObjectGUI As Long         ' 0x0C
	lObjectDefaultIIDCount As Long ' 0x10  01 00 00 00
	aObjectEventsIIDTable As Long  ' 0x14
	lObjectEventsIIDCount As Long  ' 0x18
	aObjectDefaultIIDTable As Long ' 0x1C
	ControlCount As Long           ' 0x20
	aControlArray As Long          ' 0x24
	iEventCount As Integer         ' 0x28 (40d) Number of Events
	iPCodeCount As Integer         ' 0x2C
	oInitializeEvent As Integer    ' 0x2C (44d) Offset to Initialize Event from aMethodLinkTable
	oTerminateEvent As Integer     ' 0x2E (46d) Offset to Terminate Event from aMethodLinkTable
	aEventLinkArray As Long        ' 0x30  Pointer to pointers of MethodLink
	aBasicClassObject As Long      ' 0x34 Pointer to an in-memory
	Null3 As Long                  ' 0x38
	Flag2 As Long                  ' 0x3C usually null
	*/
};

static_assert(sizeof(OptionalObjectInfo) == 0x40, "OptionalObjectInfo size is incorrect");



struct variableSizeInfo
{
	DWORD iConst1;
	DWORD iSize;
};

struct ProjectData
{
	DWORD dwVersion;							/* 5.00 in Hex (0x1F4). Version */
	struct ObjectTable* lpObjectTable;			/* Pointer to the Object Table */
	DWORD dwNull;								/* Unused value after compilation */
	LPVOID lpCodeStart;							/* Points to start of code. Unused */
	LPVOID lpCodeEnd;							/* Points to end of code. Unused */
	DWORD dwDataSize;							/* Size of VB Object Structures. Unused */
	LPVOID lpThreadSpace;						/* Pointer to Pointer to Thread Object */
	LPVOID lpVbaSeh;							/* Pointer to VBA Exception Handler */
	LPVOID lpNativeCode;						/* Pointer to .DATA section */
	char szPathInformation[527];				/* Contains Path and ID string. < SP6 */
	LPVOID lpExternalTable;						/* Pointer to External Table */
	DWORD dwExternalCount;						/* Objects in the External Table */
};
static_assert(sizeof(ProjectData) == 0x23C, "ProjectData size is incorrect");

struct ProjectInfo2
{
	LPVOID lpHeapLink;							/* Unused after compilation, always 0. */
	struct ObjectTable* lpObjectTable;			/* Back - Pointer to the Object Table. */
	DWORD dwReserved;							/* Always set to - 1 after compiling.Unused */
	DWORD dwUnused;								/* Not written or read in any case. */
	LPVOID lpObjectList;						/* Pointer to Object Descriptor Pointers. */
	DWORD dwUnused2;							/* Not written or read in any case. */
	LPSTR szProjectDescription;					/* Pointer to Project Description */
	LPSTR szProjectHelpFile;					/* Pointer to Project Help File */
	DWORD dwReserved2;							/* Always set to - 1 after compiling.Unused */
	DWORD dwHelpContextId;						/* Help Context ID set in Project Settings. */
};
static_assert(sizeof(ProjectInfo2) == 0x28, "ProjectInfo2 size is incorrect");

struct ObjectTable
{
	LPVOID lpHeapLink;							/* Unused after compilation, always 0. */
	LPVOID lpExecProj;							/* Pointer to VB Project Exec COM Object. */
	struct ProjectInfo2* lpProjectInfo2;		/* Secondary Project Information. */
	DWORD dwReserved;							/* Always set to -1 after compiling. Unused */
	DWORD dwNull;								/* Not used in compiled mode. */
	struct ProjectData* lpProjectObject;		/* Pointer to in-memory Project Data. */
	GUID uuidObject;							/* GUID of the Object Table. */
	WORD fCompileState;							/* Internal flag used during compilation. */
	WORD wTotalObjects;							/* Total objects present in Project. */
	WORD wCompiledObjects;						/* Equal to above after compiling. */
	WORD wObjectsInUse;							/* Usually equal to above after compile. */
	LPVOID lpObjectArray;						/* Pointer to Object Descriptors */
	WORD fIdeFlag;								/* Flag/Pointer used in IDE only. */
	LPVOID lpIdeData;							/* Flag/Pointer used in IDE only. */
	LPVOID lpIdeData2;							/* Flag/Pointer used in IDE only. */
	LPCSTR lpszProjectName;						/* Pointer to Project Name. */
	DWORD  dwLcid;								/* LCID of Project. */
	DWORD  dwLcid2;								/* Alternate LCID of Project. */
	LPVOID lpIdeData3;							/* Flag/Pointer used in IDE only. */
	DWORD  dwIdentifier;						/* Template Version of Structure */
};
static_assert(sizeof(ObjectTable) == 0x54, "ObjectTable size is incorrect");

struct PrivateObjectDescriptor
{
	LPVOID lpHeapLink;							/* Unused after compilation, always 0. */
	struct ObjectInfo* lpObjectInfo;			/* Pointer to the Object Info for this Object. */
	DWORD dwReserved;							/* Always set to -1 after compiling. */
	DWORD dwIdeData[3];							/* Not valid after compilation. */
	LPVOID lpObjectList;						/* Points to the Parent Structure (Array) */
	DWORD dwIdeData2;							/* Not valid after compilation. */
	LPVOID lpObjectList2[3];					/* Points to the Parent Structure (Array). */
	DWORD dwIdeData3[3];						/* Not valid after compilation. */
	DWORD dwObjectType;							/* Type of the Object described. */
	DWORD dwIdentifier;							/* Template Version of Structure. */
};
static_assert(sizeof(PrivateObjectDescriptor) == 0x40, "PrivateObjectDescriptor size is incorrect");

struct PublicObjectDescriptor
{
	struct ObjectInfo* lpObjectInfo;			/* Pointer to the Object Info for this Object */
	DWORD dwReserved;							/* Always set to -1 after compiling */
	struct variableSizeInfo* lpPublicBytes;		/* Pointer to Public Variable Size integers */
	struct variableSizeInfo* lpStaticBytes;		/* Pointer to Static Variable Size integers */
	LPVOID lpModulePublic;						/* Pointer to Public Variables in DATA section */
	LPVOID lpModuleStatic;						/* Pointer to Static Variables in DATA section */
	LPCSTR lpszObjectName;						/* Name of the Object */
	DWORD dwMethodCount;						/* Number of Methods in Object */
	LPCSTR* lpMethodNames;						/* If present, pointer to Method names array */
	DWORD bStaticVars;							/* Offset to where to copy Static Variables */
	DWORD fObjectType;							/* Flags defining the Object Type */
	DWORD dwNull;								/* Not valid after compilation */
};
static_assert(sizeof(PublicObjectDescriptor) == 0x30, "PublicObjectDescriptor size is incorrect");

struct ObjectInfo
{
	WORD wRefCount;							/* Always 1 after compilation */
	WORD wObjectIndex;						/* Index of this Object */
	struct ObjectTable* lpObjectTable;		/* Pointer to the Object Table */
	LPVOID lpIdeData;						/* Zero after compilation. Used in IDE only */
	struct PrivateObjectDescriptor* lpPrivateObject;					/* Pointer to Private Object Descriptor */
	DWORD dwReserved;						/* Always -1 after compilation */
	DWORD dwNull;							/* Unused */
	struct PublicObjectDescriptor* lpObject;/* Back-Pointer to Public Object Descriptor */
	struct ProjectData* lpProjectData;		/* Pointer to in-memory Project Object */
	WORD wMethodCount;						/* Number of Methods */
	WORD wMethodCount2;						/* Zeroed out after compilation. IDE only */
	LPVOID lpMethods;						/* Pointer to Array of Methods */
	WORD wConstants;						/* Number of Constants in Constant Pool */
	WORD wMaxConstants;						/* Constants to allocate in Constant Pool */
	LPVOID lpIdeData2;						/* Valid in IDE only */
	LPVOID lpIdeData3;						/* Valid in IDE only */
	LPVOID lpConstants;						/* Pointer to Constants Pool */
};
static_assert(sizeof(ObjectInfo) == 0x38, "ObjectInfo size is incorrect");

struct ObjectInfoWithOptional
{
	ObjectInfo	hdr;
	OptionalObjectInfo opt;
};

struct VBHeader
{
	char szVbMagic[4];							/* “VB5!” String */
	WORD wRuntimeBuild;							/* Build of the VB6 Runtime */
	char szLangDll[14];							/* Language Extension DLL */
	char szSecLangDll[14];						/* 2nd Language Extension DLL */
	WORD wRuntimeRevision;						/* Internal Runtime Revision */
	DWORD dwLCID;								/* LCID of Language DLL */
	DWORD dwSecLCID;							/* LCID of 2nd Language DLL */
	LPVOID lpSubMain;							/* Pointer to Sub Main Code */
	struct ProjectData* lpProjectData;			/* Pointer to Project Data */
	DWORD fMdlIntCtls;							/* VB Control Flags for IDs < 32 */
	DWORD fMdlIntCtls2;							/* VB Control Flags for IDs > 32 */
	DWORD dwThreadFlags;						/* Threading Mode */
	DWORD dwThreadCount;						/* Threads to support in pool */
	WORD wFormCount;							/* Number of forms present */
	WORD wExternalCount;						/* Number of external controls */
	DWORD dwThunkCount;							/* Number of thunks to create */
	LPVOID lpGuiTable;							/* Pointer to GUI Table */
	LPVOID lpExternalTable;						/* Pointer to External Table */
	struct tagREGDATA* lpComRegisterData;		/* Pointer to COM Information */
	DWORD bSZProjectDescription;				/* Offset to Project Description */
	DWORD bSZProjectExeName;					/* Offset to Project EXE Name */
	DWORD bSZProjectHelpFile;					/* Offset to Project Help File */
	DWORD bSZProjectName;						/* Offset to Project Name */
};
static_assert(sizeof(VBHeader) == 0x68, "VBHeader size is incorrect");

#pragma pack(push, 1)
struct tagREGDATA
{
	DWORD bRegInfo;								/* Offset to COM Interfaces Info (tagRegInfo) */
	DWORD bSZProjectName;						/* Offset to Project / Typelib Name */
	DWORD bSZHelpDirectory;						/* Offset to Help Directory */
	DWORD bSZProjectDescription;				/* Offset to Project Description */
	UUID uuidProjectClsId;						/* CLSID of Project / Typelib */
	DWORD dwTlbLcid;							/* LCID of Type Library */
	WORD wUnknown;								/* Might be something */
	WORD wTlbVerMajor;							/* Typelib Major Version */
	WORD wTlbVerMinor;							/* Typelib Minor Version */
};
static_assert(sizeof(tagREGDATA) == 0x2a, "tagREGDATA size is incorrect");
#pragma pack(pop)

struct tagRegInfo
{
	DWORD bNextObject;							/* Offset to COM Interfaces Info */
	DWORD bObjectName;							/* Offset to Object Name */
	DWORD bObjectDescription;					/* Offset to Object Description */
	DWORD dwInstancing;							/* Instancing Mode */
	DWORD dwObjectId;							/* Current Object ID in the Project */
	UUID uuidObject;							/* CLSID of Object */
	DWORD fIsInterface;							/* Specifies if the next CLSID is valid */
	DWORD bUuidObjectIFace;						/* Offset to CLSID of Object Interface */
	DWORD bUuidEventsIFace;						/* Offset to CLSID of Events Interface */
	DWORD fHasEvents;							/* Specifies if the CLSID above is valid */
	DWORD dwMiscStatus;							/* OLEMISC Flags(see MSDN docs) */
	__int8 fClassType;							/* Class Type */
	__int8 fObjectType;							/* Flag identifying the Object Type (VB_COM_OBJ_TYPEs) */
	WORD wToolboxBitmap32;						/* Control Bitmap ID in Toolbox */
	WORD wDefaultIcon;							/* Minimized Icon of Control Window */
	WORD fIsDesigner;							/* Specifies whether this is a Designer */
	DWORD bDesignerData;						/* Offset to Designer Data */
};
static_assert(sizeof(tagRegInfo) == 0x44, "tagRegInfo size is incorrect");


/* TODO Designer info (tagRegInfo + other fields) */


#pragma pack(push, 4)
struct struct_v5
{
	int field_0;
	HMODULE hInstance;
	int field_8;
};
#pragma pack(pop)


struct serDllTemplate
{
	char *lpLibraryNameA;
	char *lpProcAddressA;
	char *lpProcAddressW;
	struct struct_v5 *ptrStruct_v5;
};

