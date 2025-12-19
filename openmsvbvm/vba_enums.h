#pragma once

typedef enum
{
	VB_FMODE_INPUT = 0x0001,
	VB_FMODE_OUTPUT = 0x0002,
	VB_FMODE_RANDOM = 0x0004,
	VB_FMODE_BINARY = 0x0020,

	VB_FMODE_ACCESS_WRITE = 0x0200,
	VB_FMODE_ACCESS_READ = 0x0100,
} vbaFileOpenMode;

typedef enum
{
	VB_COM_OBJ_TYPE_DESIGNER = 0x2, // Designer A Visual Basic Designer for an Add - In
	VB_COM_OBJ_TYPE_CLASS = 0x10, //Class Module A Visual Basic Class
	VB_COM_OBJ_TYPE_USER_CONTROL = 0x20, //User Control A Visual Basic Active X User Control(OCX)
	VB_COM_OBJ_TYPE_USER_DOCUMENT = 0x80, //User Document A Visual Basic User Document
} VB_COM_OBJ_TYPEs;

typedef enum
{
	ApartmentModel = 0x1,			/* Specifies multi - threading using an apartment model. */
	RequireLicense = 0x2,			/* Specifies to do license validation(OCX only). */
	Unattended = 0x4,				/* Specifies that no GUI elements should be initialized. */
	SingleThreaded = 0x8,			/* Specifies that the image is single - threaded. */
	Retained = 0x10,				/* Specifies to keep the file in memory(Unattended only) */
} VB_THREAD_FLAGS;