#pragma once

#include <objbase.h>
#include <unknwn.h> // Añade esta cabecera para las anotaciones de COM
#include "vba_structures.h"

/**
 * @brief			TBD
 * @param			TBD
 * @returns			TBD
 */
EXPORT HRESULT __stdcall __vbaNew2(
	vba_new_data_arg_t* pvbNewData,
	struct IUnknown** ppv
);
