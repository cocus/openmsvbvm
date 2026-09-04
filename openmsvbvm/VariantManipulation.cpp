#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"
#include "vba_Locale.h"

#include "VariantManipulation.hpp"
#include "StringManipulation.hpp"
#include "ObjectManipulation.hpp"

#include "vba_ole_bridge_macros.h"

#define BUNCH_OF_VAR_MANIPULATIONS(declr) \
    declr(__vbaVarAdd, VarAdd) declr(__vbaVarAnd, VarAnd) declr(__vbaVarCat, VarCat) declr(__vbaVarDiv, VarDiv) \
        declr(__vbaVarEqv, VarEqv) declr(__vbaVarIdiv, VarIdiv) declr(__vbaVarImp, VarImp) declr(__vbaVarMod, VarMod) \
            declr(__vbaVarMul, VarMul) declr(__vbaVarOr, VarOr) declr(__vbaVarPow, VarPow) declr(__vbaVarSub, VarSub) \
                declr(__vbaVarXor, VarXor)

#pragma warning(disable : 4477) /* This warning is created by converting types in sprintf */

BUNCH_OF_VAR_MANIPULATIONS(DECLARE_VBA_VARIANT_MANIPULATION_BRIDGE_TO_OLE_MANIPULATION);

/**
 * @brief			Gets the value of a VT_BYREF Variant to another Variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest always.
 */
VARIANTARG* __stdcall VarDerefByref(VARIANTARG* pvargDest, VARIANTARG* pvargSrc)
{
    // Remove (mask) VT_BYREF
    pvargDest->vt = pvargSrc->vt & ~VT_BYREF;

    switch (pvargDest->vt)
    {
    case VT_I2:
    case VT_BOOL:
    {
        pvargDest->iVal = *pvargSrc->piVal;
        break;
    } /* VT_I2, VT_BOOL */

    case VT_I4:
    case VT_R4:
    case VT_BSTR:
    case VT_ERROR:
    {
        pvargDest->lVal = *pvargSrc->plVal;
        break;
    } /* VT_I4, VT_R4, VT_BSTR, VT_ERROR */

    case VT_R8:
    case VT_CY:
    case VT_DATE:
    {
        pvargDest->cyVal = *pvargSrc->pcyVal;
        break;
    } /* VT_R8, VT_CY, VT_DATE */

    case VT_VARIANT:
    case VT_DECIMAL:
    {
        pvargDest->decVal = *pvargSrc->pdecVal;
        break;
    } /* VT_VARIANT, VT_DECIMAL */

    case VT_UI1:
    {
        pvargDest->bVal = *pvargSrc->pbVal;
        break;
    } /* VT_UI1 */

    case VT_RECORD:
    {
        pvargDest->pvRecord = pvargSrc->pvRecord;
        pvargDest->pRecInfo = pvargSrc->pRecInfo;
        break;
    } /* VT_RECORD */

    case VT_DISPATCH:
    case VT_UNKNOWN:
    {
        pvargDest->plVal = (LONG*)*pvargSrc->plVal;
        break;
    } /* VT_DISPATCH, VT_UNKNOWN */

    default:
    {
        if (!(pvargSrc->vt & VT_ARRAY))
        {
            vbaRaiseException(VBA_EXCEPTION_VARIABLE_USES_A_TYPE_NOT_SUPPORTED_IN_VISUAL_BASIC);
        }

        pvargDest->lVal = *pvargSrc->plVal;
    } /* default */
    } /* switch (pvargSrc->vt) */

    return pvargDest;
} /* VarDerefByref */

/**
 * @brief			Returns a BSTR in a Variant with the hexadecimal representation of a Variant argument.
 * @param			pvargOut		Output variant variable (will be set to VT_BSTR).
 * @param			pvargIn			Input variant to convert.
 */
EXPORT void __stdcall rtcHexVarFromVar(VARIANTARG* pvargOut, VARIANTARG* pvargIn)
{
    VARIANTARG vargLocalDeRef;

    /*
        TODO: List of types:
        VT_CY	= 6,
        VT_DATE	= 7,
        VT_BSTR	= 8,
        VT_DISPATCH	= 9,
        VT_ERROR	= 10,
        VT_BOOL	= 11,
        VT_VARIANT	= 12,
        VT_UNKNOWN	= 13,
        VT_DECIMAL	= 14,
        VT_INT	= 22,
        VT_UINT	= 23,
    */

    wchar_t lpszwBuffer[20];

    // De-Ref if necessary
    if (pvargIn->vt & VT_BYREF)
    {
        VarDerefByref(&vargLocalDeRef, pvargIn);
        pvargIn = &vargLocalDeRef;
        LOG(LOG_DEBUG) << L"pvargIn->vt & VT_BYREF => deref";
    }

    /* Confirmed via real msvbvm60.dll disassembly (ordinal 573): Hex(Null) propagates
       Null rather than producing a string -- this used to fall through to the
       default "00??" case below instead. */
    if (pvargIn->vt == VT_NULL)
    {
        pvargOut->vt = VT_NULL;
        return;
    }

    // For each size of variant, format a different string
    switch (pvargIn->vt)
    {
    case VT_I1:
    case VT_UI1:
    {
        swprintf(lpszwBuffer, 19, L"%.2X", pvargIn->iVal);

        LOG(LOG_DEBUG) << L" 1 byte input '" << lpszwBuffer << L"'";

        break;
    } /* VT_I1, VT_UI1 */

    case VT_I2:
    case VT_UI2:
    {
        swprintf(lpszwBuffer, 19, L"%.4X", pvargIn->iVal);

        LOG(LOG_DEBUG) << L" 2 bytes input '" << lpszwBuffer << L"'";

        break;
    } /* VT_I2, VT_UI2 */

    case VT_I4:
    case VT_R4:
    case VT_UI4:
    {
        swprintf(lpszwBuffer, 19, L"%.8X", pvargIn->lVal);

        LOG(LOG_DEBUG) << L" 4 bytes input '" << lpszwBuffer << L"'";

        break;
    } /* VT_I4, VT_R4, VT_UI4 */

    case VT_R8:
    case VT_I8:
    case VT_UI8:
    {
        swprintf(lpszwBuffer, 19, L"%.16X", pvargIn->ulVal);

        LOG(LOG_DEBUG) << L" 8 bytes input '" << lpszwBuffer << L"'";

        break;
    } /* VT_I8, VT_R8, VT_UI8 */

    default:
    {
        LOG(LOG_DEBUG) << L"vt type not handled " << vbl::Hex((unsigned long)pvargIn->vt);

        swprintf(lpszwBuffer, 19, L"00??");
    } /* default */
    } /* switch (pvargIn->vt) */

    pvargOut->vt = VT_BSTR;
    pvargOut->bstrVal = SysAllocString(lpszwBuffer);

    if (!pvargOut->bstrVal)
    {
        vbaRaiseException(VBA_EXCEPTION_OUT_OF_STRING_SPACE);
    }
} /* rtcHexVarFromVar */

/**
 * @brief			Returns a BSTR in a Variant with the hexadecimal representation of a Variant argument.
 * @param			pvargOut		Output variant variable (will be set to VT_BSTR).
 * @param			pvargIn			Input variant to convert.
 * @returns			pvargOut always.
 */
EXPORT BSTR __stdcall rtcHexBstrFromVar(VARIANTARG* pvargIn)
{
    VARIANTARG v{};
    rtcHexVarFromVar(&v, pvargIn);
    return v.bstrVal;
} /* rtcHexBstrFromVar */

/**
 * @brief			Frees the destination variant and makes a copy of the source variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest on success, 0 on error.
 */
EXPORT VARIANTARG* __fastcall __vbaVarCopy(VARIANTARG* pvargDest, VARIANTARG* pvargSource)
{

    HRESULT result = VariantCopy(pvargDest, pvargSource);

    LOG(LOG_TRACE) << L"pvargDest " << vbl::Hex(pvargDest) << L", pvargSource " << vbl::Hex(pvargSource) << L", VariantCopy = "
                   << vbl::Hres(result);

    if (result != S_OK)
    {
        vbaRaiseException(vbaErrorFromHRESULT(result));
        return nullptr;
    }

    return pvargDest;
} /* __vbaVarCopy */

/**
 * @brief			Duplicates the contents of a Variant variable to another Variant.
 * @param			pvargDest		The destination variant.
 * @param			pvargSource		The source variant.
 * @returns			pvargDest always.
 * @remark			It will try to de-ref the value of VT_BYREF variants. It will call AddRef()
 *					if it's a VT_DISPATCH or VT_UNKNOWN.
 */
EXPORT VARIANTARG* __fastcall __vbaVarDup(VARIANTARG* pvargDest, VARIANTARG* pvargSrc)
{
    VARIANTARG vargLocalDeRef{};

    if (pvargSrc->vt & VT_BYREF)
    {
        LOG(LOG_DEBUG) << L"pvargDest " << vbl::Hex((unsigned long)pvargDest) << L", pvargSrc "
                       << vbl::Hex((unsigned long)pvargSrc) << L", VT_BYREF set, so pvargSrc is now "
                       << vbl::Hex((unsigned long)&vargLocalDeRef) << L", pvargSrc->vt " << vbl::Hex((unsigned long)pvargSrc->vt);

        VarDerefByref(&vargLocalDeRef, pvargSrc);
        pvargSrc = &vargLocalDeRef;
    }
    else
    {
        LOG(LOG_DEBUG) << L"pvargDest " << vbl::Hex((unsigned long)pvargDest) << L", pvargSrc "
                       << vbl::Hex((unsigned long)pvargSrc) << L", pvargSrc->vt " << vbl::Hex((unsigned long)pvargSrc->vt);
    }

    switch (pvargSrc->vt)
    {
    case VT_DISPATCH:
    case VT_UNKNOWN:
    {
        if (pvargSrc->pdispVal)
        {
            pvargSrc->pdispVal->AddRef();
        }
        // Fall-thru
    } /* VT_DISPATCH, VT_UNKNOWN */

    case VT_EMPTY:
    case VT_NULL:
    case VT_I2:
    case VT_I4:
    case VT_R4:
    case VT_R8:
    case VT_CY:
    case VT_DATE:
    case VT_ERROR:
    case VT_BOOL:
    case VT_DECIMAL:
    case VT_UI1:
    {
        /* Confirmed via real msvbvm60.dll disassembly (ordinal 144): VT_I1 is NOT
           in this raw-copy bucket -- it falls through to __vbaVarCopy below, same
           as everything else the jump table doesn't special-case. */
        memcpy(pvargDest, pvargSrc, sizeof(VARIANTARG));
        break;
    }

    default:
    {
        __vbaVarCopy(pvargDest, pvargSrc);
        break;
    } /* default */
    } /* switch (pvargSrc->vt) */

    return pvargDest;
} /* __vbaVarDup */

EXPORT VARIANTARG* __fastcall __vbaVarMove(VARIANTARG* pvargDest, VARIANTARG* pvargSrc)
{
    /* Free the destination VARIANTARG */
    __vbaFreeVar(pvargDest);

    if (pvargSrc->vt == VT_DISPATCH)
    {
        HRESULT hr = objIDispatchGetDefaultValue(pvargSrc->pdispVal, pvargDest);
        if (hr == S_OK)
        {
            __vbaFreeVar(pvargSrc);
        }
        else
        {
            vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
        }
    }
    else
    {
        __vbaVarDup(pvargDest, pvargSrc);
    }

    return pvargDest;
} /* __vbaVarMove */

EXPORT void __stdcall __vbaVarSetVar(VARIANTARG* pvargDest, VARIANTARG* pvargSrc)
{
    if (!pvargDest || !pvargSrc)
    {
        vbaRaiseException(VBA_EXCEPTION_INTERNAL_ERROR);
        return;
    }

    IDispatch* pidSrc = pvargSrc->pdispVal;
    VARTYPE vtSrc = pvargSrc->vt;

    if (pvargSrc->vt & VT_BYREF)
    {
        pidSrc = *pvargSrc->ppdispVal;
        vtSrc &= ~VT_BYREF;
    }

    if ((vtSrc != VT_DISPATCH) && (vtSrc != VT_UNKNOWN))
    {
        vbaRaiseException(VBA_EXCEPTION_TYPE_MISMATCH);
    }

    VARIANTARG dummy = *pvargDest;
    pvargDest->pdispVal = pidSrc;
    pvargDest->vt = vtSrc;
    __vbaFreeVar(&dummy);

    pvargSrc->vt = 0;
    pvargSrc->iVal = 0;
}

EXPORT void __fastcall __vbaVarZero(VARIANTARG* pvargVariant, VARIANTARG* pvargSrc)
{
    if (pvargVariant->vt > VT_DATE)
    {
        __vbaFreeVar(pvargVariant);
    }
}

/**
 * @brief			Frees a variant variable (including an array)
 * @param			pvargVariant	Pointer to a VARIANTARG that will be freed.
 * @returns			none.
 */
EXPORT void __fastcall __vbaFreeVar(VARIANTARG* pvargVariant)
{

    if (!pvargVariant)
    {
        LOG(LOG_WARN) << L"pvargVariant is NULL!";
        return;
    }

    LOG(LOG_TRACE) << L"pvargVariant " << vbl::Hex(pvargVariant) << L", pvargVariant->vt " << vbl::Hex(pvargVariant->vt);

    if (!(pvargVariant->vt & VT_BYREF))
    {
        /* Confirmed via real msvbvm60.dll disassembly (ordinal 131): VT_ARRAY must be
           checked BEFORE masking down to the base type. Masking with VT_TYPEMASK first
           (as this used to do, by switching on it directly) strips the VT_ARRAY bit,
           so e.g. an array of BSTRs would wrongly fall into the plain VT_BSTR case
           below and get SysFreeString()'d as if it were a scalar BSTR pointer, instead
           of being properly destroyed as a SafeArray. */
        if (pvargVariant->vt & VT_ARRAY)
        {
            if (pvargVariant->parray)
            {
                /* Confirmed: the real DLL checks the SAFEARRAY's own cLocks field
                   directly, not pvarVal->lVal (which this used to check -- that was
                   never the right field for this, per the "TODO: Check this
                   condition" this replaces). */
                if (pvargVariant->parray->cLocks)
                {
                    vbaRaiseException(VBA_EXCEPTION_ARRAY_FIXED_OR_TEMPORARILY_LOCKED);
                }
                else
                {
                    pvargVariant->parray->fFeatures &= 0xFFEFu;
                    SafeArrayDestroy(pvargVariant->parray);
                }
            }
        }
        else
        {
            switch (pvargVariant->vt & VT_TYPEMASK)
            {
            case VT_BSTR:
            {
                if (pvargVariant->bstrVal)
                {
                    SysFreeString(pvargVariant->bstrVal);
                }
                break;
            } /* VT_BSTR */

            case VT_DISPATCH:
            case VT_UNKNOWN:
            {
                if (pvargVariant->pdispVal)
                {
                    pvargVariant->pdispVal->Release();
                }
                break;
            } /* VT_DISPATCH, VT_UNKNOWN */

            case VT_EMPTY:
            case VT_NULL:
            case VT_I2:
            case VT_I4:
            case VT_R4:
            case VT_R8:
            case VT_CY:
            case VT_DATE:
            case VT_ERROR:
            case VT_BOOL:
            case VT_DECIMAL:
            case VT_UI1:
            {
                /* Confirmed via real disassembly: these own no external resource,
                   so the real DLL just zeroes vt directly without ever calling
                   VariantClear() for them. */
                break;
            }

            default:
            {
                /**
                 * Unrecognized type (e.g. VT_VARIANT) -- fall back to VariantClear
                 * (this is stated in the MSDN if we don't know what to do with it)
                 */
                VariantClear(pvargVariant);
                break;
            } /* default */
            } /* switch (pvargVariant->vt & VT_TYPEMASK) */
        }
    }
    else
    {
        /* Confirmed via real disassembly: VT_BYREF variants are just zeroed, with no
           dereference and no VariantClear call -- the pointed-to variable owns
           whatever it points to, not this Variant. */
        LOG(LOG_DEBUG) << L"pvargVariant->vt & VT_BYREF => just clearing vt, no dereference";
    }

    pvargVariant->vt = VT_EMPTY;

    LOG(LOG_TRACE) << L"pvargVariant->vt is now " << pvargVariant->vt;

} /* __vbaFreeVar */

/**
 * @brief			Frees an array of variant variables (including a arrays)
 * @param			dwCount			Count of elements.
 * @param			pvargVariant	Pointer to the first VARIANTARG element that will be freed.
 * @returns			none.
 */
EXPORT void __cdecl __vbaFreeVarList(unsigned int argCount, ...)
{
    VARIANTARG* pvargVariant;

    va_list args;
    va_start(args, argCount);

    while (argCount--)
    {
        pvargVariant = va_arg(args, VARIANTARG*);
        LOG(LOG_TRACE) << L"arg() " << vbl::Hex((unsigned long)pvargVariant) << L", argsRemaining " << vbl::Hex((unsigned long)argCount);

        __vbaFreeVar(pvargVariant);
    }

    va_end(args);
} /* __vbaFreeVarList */
