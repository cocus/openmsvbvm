#pragma once

#include "vba_internal.h"

typedef enum
{
	VBA_EXCEPTION_NO_ERROR = 0,
	VBA_EXCEPTION_RETURN_WITHOUT_GOSUB = 3,	// Return without GoSub
	VBA_EXCEPTION_INVALID_PROCEDURE_CALL = 5,	// Invalid procedure call
	VBA_EXCEPTION_OVERFLOW = 6,	// Overflow
	VBA_EXCEPTION_OUT_OF_MEMORY = 7,	// Out of memory
	VBA_EXCEPTION_SUBSCRIPT_OUT_OF_RANGE = 9,	// Subscript out of range
	VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED = 10,	// This array is fixed or temporarily locked
	VBA_EXCEPTION_DIVISION_BY_ZERO = 11,	// Division by zero
	VBA_EXCEPTION_TYPE_MISMATCH = 13,	// Type mismatch
	VBA_EXCEPTION_OUT_OF_STRING_SPACE = 14,	// Out of string space
	VBA_EXCEPTION_EXPRESSION_TOO_COMPLEX = 16,	// Expression too complex
	VBA_EXCEPTION_CANT_PERFORM_REQUESTED_OPERATION = 17,	// Can't perform requested operation
	VBA_EXCEPTION_USER_INTERRUPT_OCCURRED = 18,	// User interrupt occurred
	VBA_EXCEPTION_RESUME_WITHOUT_ERROR = 20,	// Resume without error
	VBA_EXCEPTION_OUT_OF_STACK_SPACE = 28,	// Out of stack space
	VBA_EXCEPTION_SUB_FUNCTION_OR_PROPERTY_NOT_DEFINED = 35,	// Sub, Function, or Property not defined
	VBA_EXCEPTION_TOO_MANY_DLL_APPLICATION_CLIENTS = 47,	// Too many DLL application clients
	VBA_EXCEPTION_ERROR_IN_LOADING_DLL = 48,	// Error in loading DLL
	VBA_EXCEPTION_BAD_DLL_CALLING_CONVENTION = 49,	// Bad DLL calling convention
	VBA_EXCEPTION_INTERNAL_ERROR = 51,	// Internal error
	VBA_EXCEPTION_BAD_FILENAME_OR_NUMBER = 52,	// Bad file name or number
	VBA_EXCEPTION_FILE_NOT_FOUND = 53,	// File not found
	VBA_EXCEPTION_BAD_FILE_MODE = 54,	// Bad file mode
	VBA_EXCEPTION_FILE_ALREADY_OPEN = 55,	// File already open
	VBA_EXCEPTION_DEVICE_IO_ERROR = 57,	// Device I/O error
	VBA_EXCEPTION_FILE_ALREADY_EXISTS = 58,	// File already exists
	VBA_EXCEPTION_BAD_RECORD_LENGTH = 59,	// Bad record length
	VBA_EXCEPTION_DISK_FULL = 61,	// Disk full
	VBA_EXCEPTION_INPUT_PAST_END_OF_FILE = 62,	// Input past end of file
	VBA_EXCEPTION_BAD_RECORD_NUMBER = 63,	// Bad record number
	VBA_EXCEPTION_TOO_MANY_FILES = 67,	// Too many files
	VBA_EXCEPTION_DEVICE_UNAVAILABLE = 68,	// Device unavailable
	VBA_EXCEPTION_PERMISSION_DENIED = 70,	// Permission denied
	VBA_EXCEPTION_DISK_NOT_READY = 71,	// Disk not ready
	VBA_EXCEPTION_CANT_RENAME_WITH_DIFFERENT_DRIVE = 74,	// Can't rename with different drive
	VBA_EXCEPTION_PATH_FILE_ACCESS_ERROR = 75,	// Path/File access error
	VBA_EXCEPTION_PATH_NOT_FOUND = 76,	// Path not found
	VBA_EXCEPTION_OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET = 91,	// Object variable or With block variable not set
	VBA_EXCEPTION_FOR_LOOP_NOT_INITIALIZED = 92,	// For loop not initialized
	VBA_EXCEPTION_INVALID_PATTERN_STRING = 93,	// Invalid pattern string
	VBA_EXCEPTION_INVALID_USE_OF_NULL = 94,	// Invalid use of Null
	VBA_EXCEPTION_CANT_CALL_FRIEND_PROCEDURE = 97,	// Can't call Friend procedure on an object that is not an instance of the defining class
	VBA_EXCEPTION_PROPERTY_OR_METHOD_CALL_INCLUDING_REFERENCE_TO_PRIVATE_OBJECT = 98,	// A property or method call cannot include a reference to a private object, either as an argument or as a return value
	VBA_EXCEPTION_SYSTEM_DLL_COULD_NOT_BE_LOADED = 298,	// System DLL could not be loaded
	VBA_EXCEPTION_CANT_USE_CHARACTER_DEVICES_IN_SPECIFIED_FILE_NAMES = 320,	// Can't use character device names in specified file names
	VBA_EXCEPTION_INVALID_FILE_FORMAT = 321,	// Invalid file format
	VBA_EXCEPTION_CANT_CREATE_NECESSARY_TEMPORARY_FILE = 322,	// Can’t create necessary temporary file
	VBA_EXCEPTION_INVALID_FORMAT_IN_RESOURCE_FILE = 325,	// Invalid format in resource file
	VBA_EXCEPTION_DATA_VALUE_NAMED_NOT_FOUND = 327,	// Data value named not found
	VBA_EXCEPTION_ILLEGAL_PARAMETER_CANT_WRITE_ARRAYS = 328,	// Illegal parameter; can't write arrays
	VBA_EXCEPTION_COULD_NOT_ACCESS_SYSTEM_REGISTRY = 335,	// Could not access system registry
	VBA_EXCEPTION_COMPONENT_NOT_CORRECTLY_REGISTRED = 336,	// Component not correctly registered
	VBA_EXCEPTION_COMPONENT_NOT_FOUND = 337,	// Component not found
	VBA_EXCEPTION_COMPONENT_DID_NOT_RUN_CORRECTLY = 338,	// Component did not run correctly
	VBA_EXCEPTION_OBJECT_ALREADY_LOADED = 360,	// Object already loaded
	VBA_EXCEPTION_CANT_LOAD_OR_UNLOAD_THIS_OBJECT = 361,	// Can't load or unload this object
	VBA_EXCEPTION_CONTROL_SPECIFIED_NOT_FOUND = 363,	// Control specified not found
	VBA_EXCEPTION_OBJECT_WAS_UNLOADED = 364,	// Object was unloaded
	VBA_EXCEPTION_UNABLE_TO_UNLOAD_WITHIN_THIS_CONTEXT = 365,	// Unable to unload within this context
	VBA_EXCEPTION_SPECIFIED_FILE_OUT_OF_DATE_PROGRAM_REQUIRES_LATER_VERSION = 368,	// The specified file is out of date. This program requires a later version
	VBA_EXCEPTION_SPECIFIED_OBJECT_CANT_BE_USED_AS_OWNER_FORM_FOR_SHOW = 371,	// The specified object can't be used as an owner form for Show
	VBA_EXCEPTION_INVALID_PROPERTY_VALUE = 380,	// Invalid property value
	VBA_EXCEPTION_INVALID_PROPERTY_ARRAY_INDEX = 381,	// Invalid property-array index
	VBA_EXCEPTION_PROPERTY_SET_CANT_BE_EXECUTED_AT_RUN_TIME = 382,	// Property Set can't be executed at run time
	VBA_EXCEPTION_PROPERTY_SET_CANT_BE_USED_WITH_A_READ_ONLY_PROPERTY = 383,	// Property Set can't be used with a read-only property
	VBA_EXCEPTION_NEED_PROPERTY_ARRAY_INDEX = 385,	// Need property-array index
	VBA_EXCEPTION_PROPERTY_SET_NOT_PERMITTED = 387,	// Property Set not permitted
	VBA_EXCEPTION_PROPERTY_GET_CANT_BE_EXECUTED_AT_RUN_TIME = 393,	// Property Get can't be executed at run time
	VBA_EXCEPTION_PROPERTY_GET_CANT_BE_EXECUTED_ON_WRITE_ONLY_PROPERTY = 394,	// Property Get can't be executed on write-only property
	VBA_EXCEPTION_FORM_ALREADY_DISPLAYED_CANT_SHOW_MODALLY = 400,	// Form already displayed; can't show modally
	VBA_EXCEPTION_CODE_MUST_CLOSE_TOPMOST_MODAL_FORM_FIRST = 402,	// Code must close topmost modal form first
	VBA_EXCEPTION_PERMISSION_TO_USE_OBJECT_DENIED = 419,	// Permission to use object denied
	VBA_EXCEPTION_PROPERTY_NOT_FOUND = 422,	// Property not found
	VBA_EXCEPTION_PROPERTY_OR_METHOD_NOT_FOUND = 423,	// Property or method not found
	VBA_EXCEPTION_OBJECT_REQUIRED = 424,	// Object required
	VBA_EXCEPTION_INVALID_OBJECT_USE = 425,	// Invalid object use
	VBA_EXCEPTION_COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT = 429,	// Component can't create object or return reference to this object
	VBA_EXCEPTION_CLASS_DOESNT_SUPPORT_AUTOMATION = 430,	// Class doesn't support Automation
	VBA_EXCEPTION_FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION = 432,	// File name or class name not found during Automation operation
	VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_PROPERTY_OR_METHOD = 438,	// Object doesn't support this property or method
	VBA_EXCEPTION_AUTOMATION_ERROR = 440,	// Automation error
	VBA_EXCEPTION_CONNECTION_TO_TYPE_LIBRARY_OR_OBJECT_LIBRARY_FOR_REMOTE_PROCESS_HAS_BEEN_LOST = 442,	// Connection to type library or object library for remote process has been lost
	VBA_EXCEPTION_AUTOMATION_DOESNT_HAVE_A_DEFAULT_VALUE = 443,	// Automation object doesn't have a default value
	VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_THIS_ACTION = 445,	// Object doesn't support this action
	VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_NAMED_ARGUMENTS = 446,	// Object doesn't support named arguments
	VBA_EXCEPTION_OBJECT_DOESNT_SUPPORT_CURRENT_LOCALE_SETTING = 447,	// Object doesn't support current locale setting
	VBA_EXCEPTION_NAMED_ARGUMENT_NOT_FOUND = 448,	// Named argument not found
	VBA_EXCEPTION_ARGUMENT_NOT_OPTIONAL_OR_INVALID_PROPERTY_ASSIGNMENT = 449,	// Argument not optional or invalid property assignment
	VBA_EXCEPTION_WRONG_NUMBER_OF_ARGUMENTS_OR_INVALID_PROPERTY_ASSIGNMENT = 450,	// Wrong number of arguments or invalid property assignment
	VBA_EXCEPTION_OBJECT_NOT_A_COLLECTION = 451,	// Object not a collection
	VBA_EXCEPTION_INVALID_ORDINAL = 452,	// Invalid ordinal
	VBA_EXCEPTION_SPECIFIED_NOT_FOUND = 453,	// Specified not found
	VBA_EXCEPTION_CODE_RESOURCE_NOT_FOUND = 454,	// Code resource not found
	VBA_EXCEPTION_CODE_RESOURCE_LOCK_ERROR = 455,	// Code resource lock error
	VBA_EXCEPTION_THIS_KEY_IS_ALREADY_ASSOCIATED_WITH_AN_ELEMENT_OF_THIS_COLLECTION = 457,	// This key is already associated with an element of this collection
	VBA_EXCEPTION_VARIABLE_USES_A_TYPE_NOT_SUPPORTED_IN_VISUAL_BASIC = 458,	// Variable uses a type not supported in Visual Basic
	VBA_EXCEPTION_THIS_COMPONENT_DOESNT_SUPPORT_THE_SET_OF_EVENTS = 459,	// This component doesn't support the set of events
	VBA_EXCEPTION_INVALID_CLIPBOARD_FORMAT = 460,	// Invalid Clipboard format
	VBA_EXCEPTION_METHOD_OR_DATA_MEMBER_NOT_FOUND = 461,	// Method or data member not found
	VBA_EXCEPTION_THE_REMOTE_SERVER_MACHINE_DOES_NOT_EXIST_OR_IS_UNAVAILABLE = 462,	// The remote server machine does not exist or is unavailable
	VBA_EXCEPTION_CLASS_NOT_REGISTREED_ON_LOCAL_MACHINE = 463,	// Class not registered on local machine
	VBA_EXCEPTION_CANT_CREATE_AUTOREDRAW_IMAGE = 480,	// Can't create AutoRedraw image
	VBA_EXCEPTION_INVALID_PICTURE = 481,	// Invalid picture
	VBA_EXCEPTION_PRINTER_ERROR = 482,	// Printer error
	VBA_EXCEPTION_PRINTER_DRIVER_DOES_NOT_SUPPORT_SPECIFIED_PROPERTY = 483,	// Printer driver does not support specified property
	VBA_EXCEPTION_PROBLEM_GETTING_PRINTER_INFORMATION_FROM_SYSTEM = 484,	// Problem getting printer information from the system. Make sure the printer is set up correctly
	VBA_EXCEPTION_INVALID_PICTURE_TYPE = 485,	// Invalid picture type
	VBA_EXCEPTION_CANT_PRINT_FORM_IMAGE_TO_THIS_TYPE_OF_PRINTER = 486,	// Can't print form image to this type of printer
	VBA_EXCEPTION_CANT_EMPTY_CLIPBOARD = 520,	// Can't empty Clipboard
	VBA_EXCEPTION_CANT_OPEN_CLIPBOARD = 521,	// Can't open Clipboard
	VBA_EXCEPTION_CANT_SAVE_FILE_TO_TEMP_DIRECTORY = 735,	// Can't save file to TEMP directory
	VBA_EXCEPTION_SEARCH_TEXT_NOT_fOUND = 744,	// Search text not found
	VBA_EXCEPTION_REPLACEMENTS_TOO_LONG = 746,	// Replacements too long
	VBA_EXCEPTION_OUT_OF_MEMORY2 = 31001,	// Out of memory
	VBA_EXCEPTION_NO_OBJECT = 31004,	// No object
	VBA_EXCEPTION_CLASS_IS_NOT_SET = 31018,	// Class is not set
	VBA_EXCEPTION_UNABLE_TO_ACTIVATE_OBJECT = 31027,	// Unable to activate object
	VBA_EXCEPTION_UNABLE_TO_CREATE_EMBEDDED_OBJECT = 31032,	// Unable to create embedded object
	VBA_EXCEPTION_ERROR_SAVING_TO_FILE = 31036,	// Error saving to file
	VBA_EXCEPTION_ERROR_LOADING_FROM_FILE = 31037,	// Error loading from file

} vbaExceptions;


void __stdcall RaiseExceptionIfLastErrorIsSet();

CEXTERN vbaExceptions __stdcall vbaErrorFromHRESULT(
	HRESULT hr
);

CEXTERN void vbaRaiseException(
	vbaExceptions exceptionCode
);