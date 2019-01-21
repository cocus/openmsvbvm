#pragma once

//#include <windows.h>




struct TObjectInfo
{
	__int16 iConst1;
	__int16 iObjectIndex;
	struct ObjectTable * aObjectTable;
	int lNull1;
	int aObjectDescriptor;
	int lConst2;
	int lNull2;
	struct TObject * aObjectHeader;
	int aObjectData;
	int lMethodCount;
	int aProcTable;
	__int16 iConstantsCount;
	__int16 iMaxConstants;
	int Flag5;
	__int16 Flag6;
	__int16 Flag7;
	int aConstantPool;
};


struct MethodLinkNative
{
	char	jmpOpCode;
	int		jmpOffset;
};

struct TObjectInfoWithOptionals
{
	TObjectInfo	hdr;

	int fDesigner;
	CLSID	*clsidObjectClass;
	CLSID	*clsidObjectInterface;
	int aGuidObjectGUI;
	int lObjectDefaultIIDCount;
	int aObjectEventsIIDTable;
	int lObjectEventsIIDCount;
	int aObjectDefaultIIDTable;
	int lControlCount;
	int aControlArray;
	__int16 iEventCount;
	__int16 iPCodeCount;
	__int16 oInitializeEvent;
	__int16 oTerminateEvent;
	void * aEventLinkArray;
	int aBasicClassObject;
	int lNull3;
	int lFlag2;
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



struct variableSizeInfo
{
	unsigned __int16 iConst1;
	unsigned __int16 iSize;
};

struct TObject
{
	struct TObjectInfo * aObjectInfo; /* pointer to TObjectInfo */
	int lConst1;
	struct variableSizeInfo * aPublicBytes;
	struct variableSizeInfo * aStaticBytes;
	int aModulePublic;
	int aModuleStatic;
	int aNTSObjectName;
	int lMethodCount;
	int aMethodNameTable;
	int oStaticVars;
	int lObjectType;
	int lNull2;
};

struct ObjectTable
{
	int lNull1;
	int aExecProj;
	int aProjectInfo2;
	int lConst1;
	int lNull2;
	int aProjectObject;
	int uuidObjectTable;
	int Flag2;
	int Flag3;
	int Flag4;
	__int16 fCompileType;
	__int16 iObjectsCount;
	__int16 iCompiledObjects;
	__int16 iObjectsInUse;
	int aObjectsArray; /* pointer to TObject */
	int lNull3;
	int lNull4;
	int lNull5;
	int aNTSProjectName;
	int lLcID1;
	int lLcID2;
	int lNull6;
	int lTemplateVersion;
};

struct ProjectInfo
{
	int		lTemplateVersion;
	int		aObjectTable; /* pointer to ObjectTable */
	int		lNull1;
	int		aStartOfCode;
	int		aEndOfCode;
	int		lDataBufferSize;
	int		aThreadSpace;
	int		aVBAExceptionhandler;
	int		aNativeCode;
	char	uIncludeID[527];
	int		aExternalTable;
	int		lExternalCount;
};

struct VBHeader
{
	char szVbMagic[4];
	__int16 wRuntimeBuild;
	char szLangDll[14];
	char szSecLangDll[14];
	__int16 wRuntimeRevision;
	int dwLCID;
	int dwSecLCID;
	int lpSubMain;
	int lpProjectData; /* pointer to ProjectInfo */
	int fMdlIntCtls;
	int fMdlIntCtls2;
	int dwThreadFlags;
	int dwThreadCount;
	__int16 wFormCount;
	__int16 wExternalCount;
	int dwThunkCount;
	int lpGuiTable;
	int lpExternalTable;
	int lpComRegisterData;
	int bSZProjectDescription;
	int bSZProjectExeName;
	int bSZProjectHelpFile;
	int bSZProjectName;
};

struct __declspec(align(4)) struct_v5
{
	int field_0;
	HMODULE hInstance;
	int field_8;
};


struct serDllTemplate
{
	char *lpLibraryNameA;
	char *lpProcAddressA;
	char *lpProcAddressW;
	struct struct_v5 *ptrStruct_v5;
};

