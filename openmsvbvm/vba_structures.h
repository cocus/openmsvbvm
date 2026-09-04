#pragma once

// #include <windows.h>

struct MethodLinkNative
{
    char jmpOpCode;
    int jmpOffset;
};

/**
 * Describes one control placed on a Form (or the Form itself -- see dwIndex below),
 * confirmed live via IDA against two independently-compiled calibration builds (one
 * with a CommandButton named/captioned identically, one with different Name/Caption/
 * Left/Top/Width/Height, to rule out coincidence) and cross-checked against Alex
 * Ionescu's "Visual Basic Image Internal Structure Format" doc's "Control Info"
 * table (http://sandsprite.com/vb-reversing/, see VB6-REVERSING-REFERENCES.md),
 * which gave the initial offsets this confirmed.
 *
 * OptionalObjectInfo.lpControls is an INLINE array of dwControlCount of these
 * (confirmed the same way ObjectTable.lpObjectArray turned out to be an inline
 * PublicObjectDescriptor array, not a pointer array) -- one entry per placed control,
 * PLUS one extra entry representing the Form/container itself: that entry's dwIndex
 * is 0xFFFFFFFF (confirmed live, twice) where every real placed control has a real
 * (small) index instead -- see FindOwnFormControlInfo (EventDispatch.hpp).
 *
 * lpEventTable is what actually matters for event dispatch: a fixed-shape table --
 * slot 0 null, slot 1 a self-reference back to this same ControlInfo, slot 2 the
 * owning class's ObjectInfoWithOptional*, slots 3-5 the real EVENT_SINK_QueryInterface/
 * AddRef/Release exports (the exact same triple FindEventSinkBlock's
 * IsRealEventSinkBlock already checks for a WithEvents sink) -- then, starting at
 * slot 6, one thunk pointer per intrinsic event THAT CONTROL TYPE (or the Form
 * itself) supports, at a FIXED ordinal matching that type's real "XxxEvents"
 * interface's own declaration order (e.g. real FormEvents: Load=6th member -> slot
 * 6+6=12; QueryUnload=10th member -> slot 6+9=15; real CommandButtonEvents: Click is
 * its 1st member -> slot 6+0=6) -- confirmed live for both Form's own Load/
 * QueryUnload and a CommandButton's Click. Null where that event isn't implemented.
 * Unlike a WithEvents sink block's per-dispId thunk array (GetHandlerThunk), these
 * slots point DIRECTLY at each thunk's start -- no extra offset correction needed
 * (confirmed live: disassembly shows the thunk's own "sub esp+4,imm" as the very
 * first instruction at the stored address). See GetFixedEventThunk (EventDispatch.hpp).
 */
struct ControlInfo
{
    DWORD fControlType; /* Type of control -- confirmed so far: CommandButton=0x00110040, Label=0x00120040 (see ControlWindow.hpp). */
    WORD wEventCount; /* Number of event names that control type's XxxEvents interface declares (not how many are implemented). */
    WORD bWEventsOffset; /* Unconfirmed -- always seen 0 so far. */
    LPVOID lpGuid; /* Pointer to this control instance's GUID (freshly randomized per compile, like other per-object GUIDs in this project). */
    DWORD dwIndex; /* Control index -- 0xFFFFFFFF for the Form/container's own entry, a real small index for a placed control. */
    DWORD dwNull;  /* Unconfirmed -- always seen 0 so far. */
    DWORD dwNull2; /* Unconfirmed -- always seen 0 so far. */
    LPVOID lpEventTable; /* Pointer to the fixed-shape event thunk table -- see this struct's own doc comment above. */
    LPVOID lpIdeData;    /* IDE only. */
    LPCSTR lpszName;     /* Name of this control (or the Form's own class name for its own entry). */
    DWORD dwIndexCopy;   /* Same value as dwIndex in every sample seen so far. */
};
static_assert(sizeof(ControlInfo) == 0x28, "ControlInfo size is incorrect");

struct OptionalObjectInfo
{
    DWORD dwObjectGuids;            /* How many GUIDs to Register. 2 = Designer */
    CLSID* clsidObjectClass;        /* Unique GUID of the Object */
    CLSID* clsidObjectInterface;    /* Unused */
    LPVOID lpuuidObjectTypes;       /* Pointer to Array of Object Interface GUIDs */
    DWORD dwObjectTypeGuids;        /* How many GUIDs in the Array above */
    LPVOID lpControls2;             /* Usually the same as lpControls */
    DWORD dwNull2;                  /* Unused */
    LPVOID lpObjectGuid2;           /* Pointer to Array of Object GUIDs */
    DWORD dwControlCount;           /* Number of Controls in array below */
    struct ControlInfo* lpControls; /* Inline array of dwControlCount ControlInfo -- see that struct's own doc comment */
    WORD wEventCount;               /* Number of Events in Event Array */
    WORD wPCodeCount;               /* Number of P-Codes used by this Object */
    WORD bWInitializeEvent;         /* Offset to Initialize Event from Event Table */
    WORD bWTerminateEvent;          /* Offset to Terminate Event in Event Table */
    LPVOID lpEvents;                /* Pointer to Events Array */
    LPVOID lpBasicClassObject;      /* Pointer to in-memory Class Objects */
    DWORD dwNull3;                  /* Unused */
    LPVOID lpIdeData;               /* Only valid in IDE */
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
    DWORD dwVersion;                   /* 5.00 in Hex (0x1F4). Version */
    struct ObjectTable* lpObjectTable; /* Pointer to the Object Table */
    DWORD dwNull;                      /* Unused value after compilation */
    LPVOID lpCodeStart;                /* Points to start of code. Unused */
    LPVOID lpCodeEnd;                  /* Points to end of code. Unused */
    DWORD dwDataSize;                  /* Size of VB Object Structures. Unused */
    LPVOID lpThreadSpace;              /* Pointer to Pointer to Thread Object */
    LPVOID lpVbaSeh;                   /* Pointer to VBA Exception Handler */
    LPVOID lpNativeCode;               /* Pointer to .DATA section */
    char szPathInformation[527];       /* Contains Path and ID string. < SP6 */
    LPVOID lpExternalTable;            /* Pointer to External Table */
    DWORD dwExternalCount;             /* Objects in the External Table */
};
static_assert(sizeof(ProjectData) == 0x23C, "ProjectData size is incorrect");

struct ProjectInfo2
{
    LPVOID lpHeapLink;                 /* Unused after compilation, always 0. */
    struct ObjectTable* lpObjectTable; /* Back - Pointer to the Object Table. */
    DWORD dwReserved;                  /* Always set to - 1 after compiling.Unused */
    DWORD dwUnused;                    /* Not written or read in any case. */
    LPVOID lpObjectList;               /* Pointer to Object Descriptor Pointers. */
    DWORD dwUnused2;                   /* Not written or read in any case. */
    LPSTR szProjectDescription;        /* Pointer to Project Description */
    LPSTR szProjectHelpFile;           /* Pointer to Project Help File */
    DWORD dwReserved2;                 /* Always set to - 1 after compiling.Unused */
    DWORD dwHelpContextId;             /* Help Context ID set in Project Settings. */
};
static_assert(sizeof(ProjectInfo2) == 0x28, "ProjectInfo2 size is incorrect");

/**
 * lpObjectArray (below) is genuinely walkable from the compiled EXE's own VBHeader --
 * confirmed live by reading it out of this project's own Debug\Proyecto1.exe (a
 * VBHeader* is already captured at startup as g_pvbhGlobal, see dllmain.cpp):
 * VBHeader.lpProjectData -> ProjectData.lpObjectTable -> here -> lpObjectArray, for
 * wObjectsInUse entries. This is NOT an array of pointers (despite Alex Ionescu's "VB
 * Image Internal Structure Format" doc calling it "Pointer to Object Descriptors",
 * ambiguous on this point) -- it's PublicObjectDescriptor structs stored INLINE,
 * contiguous, each exactly sizeof(PublicObjectDescriptor) (0x30) bytes apart; dumping
 * a real 5-object project's array byte-for-byte is what proved this. VBHeader
 * .wFormCount separately counts just the Form/MDIForm entries within this same array
 * (see PublicObjectDescriptor.fObjectType's doc comment for how those are told apart
 * from Class/Module entries).
 */
struct ObjectTable
{
    LPVOID lpHeapLink;                   /* Unused after compilation, always 0. */
    LPVOID lpExecProj;                   /* Pointer to VB Project Exec COM Object. */
    struct ProjectInfo2* lpProjectInfo2; /* Secondary Project Information. */
    DWORD dwReserved;                    /* Always set to -1 after compiling. Unused */
    DWORD dwNull;                        /* Not used in compiled mode. */
    struct ProjectData* lpProjectObject; /* Pointer to in-memory Project Data. */
    GUID uuidObject;                     /* GUID of the Object Table. */
    WORD fCompileState;                  /* Internal flag used during compilation. */
    WORD wTotalObjects;                  /* Total objects present in Project. */
    WORD wCompiledObjects;               /* Equal to above after compiling. */
    WORD wObjectsInUse;                  /* Usually equal to above after compile. */
    struct PublicObjectDescriptor* lpObjectArray; /* Inline array of wObjectsInUse PublicObjectDescriptors -- see this struct's own doc comment above. */
    WORD fIdeFlag;                                /* Flag/Pointer used in IDE only. */
    LPVOID lpIdeData;                             /* Flag/Pointer used in IDE only. */
    LPVOID lpIdeData2;                            /* Flag/Pointer used in IDE only. */
    LPCSTR lpszProjectName; /* Pointer to Project Name. */
    DWORD dwLcid;           /* LCID of Project. */
    DWORD dwLcid2;          /* Alternate LCID of Project. */
    LPVOID lpIdeData3;      /* Flag/Pointer used in IDE only. */
    DWORD dwIdentifier;     /* Template Version of Structure */
};
static_assert(sizeof(ObjectTable) == 0x54, "ObjectTable size is incorrect");

struct PrivateObjectDescriptor
{
    LPVOID lpHeapLink;               /* Unused after compilation, always 0. */
    struct ObjectInfo* lpObjectInfo; /* Pointer to the Object Info for this Object. */
    DWORD dwReserved;                /* Always set to -1 after compiling. */
    DWORD dwIdeData[3];              /* Not valid after compilation. */
    LPVOID lpObjectList;             /* Points to the Parent Structure (Array) */
    DWORD dwIdeData2;                /* Not valid after compilation. */
    LPVOID lpObjectList2[3];         /* Points to the Parent Structure (Array). */
    DWORD dwIdeData3[3];             /* Not valid after compilation. */
    DWORD dwObjectType;              /* Type of the Object described. */
    DWORD dwIdentifier;              /* Template Version of Structure. */
};
static_assert(sizeof(PrivateObjectDescriptor) == 0x40, "PrivateObjectDescriptor size is incorrect");

struct PublicObjectDescriptor
{
    struct ObjectInfo* lpObjectInfo;        /* Pointer to the Object Info for this Object */
    DWORD dwReserved;                       /* Always set to -1 after compiling */
    struct variableSizeInfo* lpPublicBytes; /* Pointer to Public Variable Size integers */
    struct variableSizeInfo* lpStaticBytes; /* Pointer to Static Variable Size integers */
    LPVOID lpModulePublic;                  /* Pointer to Public Variables in DATA section */
    LPVOID lpModuleStatic;                  /* Pointer to Static Variables in DATA section */
    LPCSTR lpszObjectName;                  /* Name of the Object */
    DWORD dwMethodCount;                    /* Number of Methods in Object */
    LPCSTR* lpMethodNames;                  /* If present, pointer to Method names array */
    DWORD bStaticVars;                      /* Offset to where to copy Static Variables */
    /**
     * Flags defining the Object Type -- a per-kind bitmask, not a small sequential
     * enum. Confirmed live (not from any external doc) by dumping this field for
     * every entry of a real 5-object project's ObjectTable.lpObjectArray (1 Module, 3
     * Class Modules, 1 Form) plus a separate calibration build with just an MDIForm:
     *   Module        0x018001  (bits 0, 15, 16)
     *   Class Module  0x118003  (bits 0, 1, 15, 16, 20)
     *   Form/MDIForm  0x018083  (bits 0, 1, 7, 15, 16)
     * Bits 0/15/16 are common to every kind seen so far (some generic "compiled"/
     * valid-descriptor markers, not type-discriminating). Bit 1 is set for both
     * instantiable kinds (Class, Form) and clear for Module (which can't be `New`'d)
     * -- likely a generic "is a class-like/instantiable object" flag. The
     * kind-specific discriminator bits confirmed so far: bit 7 (0x80) = Form or
     * MDIForm (this project's IsFormLikeDescriptor, ObjectManipulation.cpp, tests
     * exactly this bit -- MDIForm sharing it with Form was the calibration's main new
     * finding), bit 20 (0x100000) = plain Class Module. UserControl/UserDocument/
     * PropertyPage haven't been calibrated (no test project built for them yet) so
     * their bit(s) are still unknown -- compile a minimal project containing one and
     * dump this field the same way to extend this table.
     */
    DWORD fObjectType;
    DWORD dwNull; /* Not valid after compilation */
};
static_assert(sizeof(PublicObjectDescriptor) == 0x30, "PublicObjectDescriptor size is incorrect");

struct ObjectInfo
{
    WORD wRefCount;                                  /* Always 1 after compilation */
    WORD wObjectIndex;                               /* Index of this Object */
    struct ObjectTable* lpObjectTable;               /* Pointer to the Object Table */
    LPVOID lpIdeData;                                /* Zero after compilation. Used in IDE only */
    struct PrivateObjectDescriptor* lpPrivateObject; /* Pointer to Private Object Descriptor */
    DWORD dwReserved;                                /* Always -1 after compilation */
    DWORD dwNull;                                    /* Unused */
    struct PublicObjectDescriptor* lpObject;         /* Back-Pointer to Public Object Descriptor */
    struct ProjectData* lpProjectData;               /* Pointer to in-memory Project Object */
    WORD wMethodCount;                               /* Number of Methods */
    WORD wMethodCount2;                              /* Zeroed out after compilation. IDE only */
    LPVOID lpMethods;                                /* Pointer to Array of Methods */
    WORD wConstants;                                 /* Number of Constants in Constant Pool */
    WORD wMaxConstants;                              /* Constants to allocate in Constant Pool */
    LPVOID lpIdeData2;                               /* Valid in IDE only */
    LPVOID lpIdeData3;                               /* Valid in IDE only */
    LPVOID lpConstants;                              /* Pointer to Constants Pool */
};
static_assert(sizeof(ObjectInfo) == 0x38, "ObjectInfo size is incorrect");

struct ObjectInfoWithOptional
{
    ObjectInfo hdr;
    OptionalObjectInfo opt;
};

struct VBHeader
{
    char szVbMagic[4];                 /* VB5! String */
    WORD wRuntimeBuild;                /* Build of the VB6 Runtime */
    char szLangDll[14];                /* Language Extension DLL */
    char szSecLangDll[14];             /* 2nd Language Extension DLL */
    WORD wRuntimeRevision;             /* Internal Runtime Revision */
    DWORD dwLCID;                      /* LCID of Language DLL */
    DWORD dwSecLCID;                   /* LCID of 2nd Language DLL */
    LPVOID lpSubMain;                  /* Pointer to Sub Main Code */
    struct ProjectData* lpProjectData; /* Pointer to Project Data */
    DWORD fMdlIntCtls;                 /* VB Control Flags for IDs < 32 */
    DWORD fMdlIntCtls2;                /* VB Control Flags for IDs > 32 */
    DWORD dwThreadFlags;               /* Threading Mode */
    DWORD dwThreadCount;               /* Threads to support in pool */
    WORD wFormCount;                   /* Number of forms present */
    WORD wExternalCount;               /* Number of external controls */
    DWORD dwThunkCount;                /* Number of thunks to create */
    /* Checked live whether this is a "list of this project's Forms" (as a name like
       "GUI Table" suggests, and as this project briefly hoped when it confirmed
       ObjectTable.lpObjectArray is real/walkable -- see that struct's doc comment):
       it is NOT. In a real compiled EXE this instead points to a small fixed record
       -- a leading DWORD (seen live as 0x50) followed by what reads as a 16-byte GUID,
       then mostly zero bytes -- with no pointer anywhere in it back to any Form's
       PublicObjectDescriptor. Whatever this table actually is (a guess: per-project
       GUI/toolbox resource metadata, unrelated to enumerating Forms), it is NOT the
       way to enumerate a project's Forms -- that has to come from walking
       ObjectTable.lpObjectArray and testing PublicObjectDescriptor.fObjectType's
       Form/MDIForm bit (0x80) on each entry, cross-checked against wFormCount above. */
    LPVOID lpGuiTable;                    /* Pointer to GUI Table */
    LPVOID lpExternalTable;               /* Pointer to External Table */
    struct tagREGDATA* lpComRegisterData; /* Pointer to COM Information */
    DWORD bSZProjectDescription;          /* Offset to Project Description */
    DWORD bSZProjectExeName;              /* Offset to Project EXE Name */
    DWORD bSZProjectHelpFile;             /* Offset to Project Help File */
    DWORD bSZProjectName;                 /* Offset to Project Name */
};
static_assert(sizeof(VBHeader) == 0x68, "VBHeader size is incorrect");

#pragma pack(push, 1)
struct tagREGDATA
{
    DWORD bRegInfo;              /* Offset to COM Interfaces Info (tagRegInfo) */
    DWORD bSZProjectName;        /* Offset to Project / Typelib Name */
    DWORD bSZHelpDirectory;      /* Offset to Help Directory */
    DWORD bSZProjectDescription; /* Offset to Project Description */
    UUID uuidProjectClsId;       /* CLSID of Project / Typelib */
    DWORD dwTlbLcid;             /* LCID of Type Library */
    WORD wUnknown;               /* Might be something */
    WORD wTlbVerMajor;           /* Typelib Major Version */
    WORD wTlbVerMinor;           /* Typelib Minor Version */
};
static_assert(sizeof(tagREGDATA) == 0x2a, "tagREGDATA size is incorrect");
#pragma pack(pop)

struct tagRegInfo
{
    DWORD bNextObject;        /* Offset to COM Interfaces Info */
    DWORD bObjectName;        /* Offset to Object Name */
    DWORD bObjectDescription; /* Offset to Object Description */
    DWORD dwInstancing;       /* Instancing Mode */
    DWORD dwObjectId;         /* Current Object ID in the Project */
    UUID uuidObject;          /* CLSID of Object */
    DWORD fIsInterface;       /* Specifies if the next CLSID is valid */
    DWORD bUuidObjectIFace;   /* Offset to CLSID of Object Interface */
    DWORD bUuidEventsIFace;   /* Offset to CLSID of Events Interface */
    DWORD fHasEvents;         /* Specifies if the CLSID above is valid */
    DWORD dwMiscStatus;       /* OLEMISC Flags(see MSDN docs) */
    __int8 fClassType;        /* Class Type */
    __int8 fObjectType;       /* Flag identifying the Object Type (VB_COM_OBJ_TYPEs) */
    WORD wToolboxBitmap32;    /* Control Bitmap ID in Toolbox */
    WORD wDefaultIcon;        /* Minimized Icon of Control Window */
    WORD fIsDesigner;         /* Specifies whether this is a Designer */
    DWORD bDesignerData;      /* Offset to Designer Data */
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
    char* lpLibraryNameA;
    char* lpProcAddressA;
    char* lpProcAddressW;
    struct struct_v5* ptrStruct_v5;
};

typedef struct
{
    DWORD flags;
    LPCLSID lpguidCoClass;
    LPGUID lpguidInterface;
    unsigned int dummy2;
} vba_new_data_arg_t;
static_assert(sizeof(vba_new_data_arg_t) == 0x10, "vba_new_data_arg_t size is incorrect");
