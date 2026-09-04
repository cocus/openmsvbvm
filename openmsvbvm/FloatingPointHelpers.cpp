#include "vba_internal.h"
#include "Logging.hpp"
#include "Exceptions.hpp"

#include <cmath>
#include <cerrno>

#define HIDWORD(d) (d >> 16)

EXPORT __int64 __stdcall __allmul(__int64 i64A, __int64 i64B)
{
    __int64 result;

    if (HIDWORD(i64A) | HIDWORD(i64B))
        result = i64A * i64B;
    else
        result = (unsigned int)i64B * (unsigned __int64)(unsigned int)i64A;
    return result;
}

EXPORT double __stdcall __vbaCyMul(double dblA, double dblB)
{
    return dblA * dblB / 1000.0; /* Because VB treats the Currencies as 1000 times their value */
}

EXPORT unsigned char __fastcall __vbaUI1I2(unsigned __int16 uiIn)
{
    if (uiIn > 0xFFu)
    {
        vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
    }

    return (unsigned char)uiIn;
}

/**
 * @brief			Narrows a Long (I4) to a Byte (UI1), raising Overflow if it
 *					doesn't fit in 0..255.
 */
EXPORT unsigned char __fastcall __vbaUI1I4(unsigned long ulIn)
{
    if (ulIn > 0xFFu)
    {
        vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
    }

    return (unsigned char)ulIn;
}

EXPORT __declspec(naked) void _adj_fptan(void)
{
    _asm
    {
		fptan
		retn
    }
}

EXPORT __declspec(naked) void _adj_fpatan(void)
{
    _asm
    {
		fpatan
		retn
    }
}

/*
 * Confirmed via real msvbvm60.dll disassembly: every _CIxxx function below is the
 * classic MSVC "checked FPU instruction" thunk. They take NO conventional arguments --
 * the value arrives on the x87 register stack (ST(0)), exactly like the raw FPU
 * instruction they wrap, and compiled code calls them as `fld value; call _CIxxx;
 * fstp result`. The previous versions here were declared `__stdcall(double value)`,
 * which is a completely different (and wrong) calling convention -- they read a
 * nonexistent stack argument and never touched ST(0) at all, so every VB6 Sin/Cos/Tan/
 * Atn/Log/Sqr/Exp call was silently broken. Fixed as __declspec(naked) functions that
 * operate on the FPU stack directly, matching the already-correct _adj_fptan/
 * _adj_fpatan below (which somehow already got this right while these didn't).
 * Domain-error checking (e.g. Sqr of a negative number, real DLL raises "Invalid
 * procedure call") is NOT replicated here -- out-of-scope for this pass; invalid
 * domains silently produce IEEE754 NaN/Inf instead of a VB runtime error.
 */

EXPORT __declspec(naked) void _CIsin(void)
{
    _asm
    {
		fsin
		retn
    }
}

EXPORT __declspec(naked) void _CIcos(void)
{
    _asm
    {
		fcos
		retn
    }
}

EXPORT __declspec(naked) void _CIsqrt(void)
{
    _asm
    {
		fsqrt
		retn
    }
}

EXPORT __declspec(naked) void _CIatan(void)
{
    /* Confirmed: real DLL computes atan(x) via FPATAN (which computes atan2(ST(1),
       ST(0))) by loading 1.0 onto ST(0) first, so it becomes atan(x/1) = atan(x). */
    _asm
    {
		fld1
		fpatan
		retn
    }
}

EXPORT __declspec(naked) void _CItan(void)
{
    /* Confirmed: FPTAN computes tan(ST(0)) but also pushes an extra 1.0 onto the FPU
       stack afterward -- the real DLL immediately pops that extra value off. Not
       replicated here: the real DLL's fallback argument-reduction path (via FPREM1)
       for arguments too large for FPTAN to handle directly. */
    _asm
    {
		fptan
		fstp st(0)
		retn
    }
}

EXPORT __declspec(naked) void _CIlog(void)
{
    /* Confirmed: natural log via FYL2X (computes ST(1)*log2(ST(0)), pops both, pushes
       result) -- load ln(2) onto ST(0) then swap so this computes ln(2)*log2(x) =
       ln(x). */
    _asm
    {
		fldln2
		fxch st(1)
		fyl2x
		retn
    }
}

static double __cdecl _CIexp_impl(double value)
{
    return exp(value);
}

EXPORT __declspec(naked) void _CIexp(void)
{
    /* The real DLL's own exp() implementation goes several helper functions deep
       (control-word save/restore plus its own internal exp algorithm) and wasn't
       fully traced -- this achieves the same FPU-stack-neutral (1 value in, 1 value
       out) contract using the C library's exp() instead of porting that algorithm. */
    _asm
    {
		sub esp, 8
		fstp qword ptr [esp]
		call _CIexp_impl
		add esp, 8
		retn
    }
}

/**
 * @brief			Rounds ST(0) to the nearest 16-bit integer (current FPU rounding
 *					mode) and returns it in EAX, raising Overflow if it doesn't fit.
 *					Confirmed via real msvbvm60.dll disassembly (ordinal 124).
 */
EXPORT __declspec(naked) void __vbaFpI2(void)
{
    _asm
    {
		sub esp, 4
		fistp word ptr [esp]
		fnstsw ax
		test al, 1
		jnz overflow
		movsx eax, word ptr [esp]
		add esp, 4
		retn
	overflow:
		add esp, 4
		push 6
		call vbaRaiseException
		add esp, 4
		xor eax, eax
		retn
    }
}

/*
 * Confirmed via real msvbvm60.dll disassembly: every _adj_fdiv_ and _adj_fdivr_
 * function below exists ONLY to work around the Pentium FDIV erratum (a hardware division bug
 * in early-generation Pentium CPUs from 1994) -- the real DLL checks whether the
 * operands fall in the narrow "affected" range and only then takes a slow corrected
 * path; otherwise it does a plain FDIV. On every CPU that could plausibly run this
 * DLL today, the erratum simply does not exist, so the correct AND simpler modern
 * implementation is to just perform the division directly and unconditionally --
 * this gives the exact right answer with no need to detect or work around a bug that
 * isn't there. Each takes the divisor by value on the stack (the dividend is already
 * on the FPU stack, ST(0)); the "r" (reversed) variants divide the OTHER way
 * (divisor / ST(0)), matching FDIVR/FIDIVR.
 */

EXPORT __declspec(naked) void _adj_fdiv_m16i(void)
{
    _asm
        {
		fidiv word ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdiv_m32(void)
{
    _asm
        {
		fdiv dword ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdiv_m32i(void)
{
    _asm
        {
		fidiv dword ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdiv_m64(void)
{
    _asm
        {
		fdiv qword ptr [esp+4]
		retn 8
        }
}

EXPORT __declspec(naked) void _adj_fdivr_m16i(void)
{
    _asm
        {
		fidivr word ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdivr_m32(void)
{
    _asm
        {
		fdivr dword ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdivr_m32i(void)
{
    _asm
        {
		fidivr dword ptr [esp+4]
		retn 4
        }
}

EXPORT __declspec(naked) void _adj_fdivr_m64(void)
{
    _asm
        {
		fdivr qword ptr [esp+4]
		retn 8
        }
}

/*
 * FPREM/FPREM1 could only complete a partial reduction per instruction on the
 * original 8087/80287 FPUs, requiring software to loop on the C2 flag until it
 * cleared. Every 80387 and later (i.e. anything since 1987) always completes in one
 * instruction, so the loop below is dead code on any real machine -- kept only as a
 * correctness safety net, matching the exact idiom the real DLL itself uses elsewhere
 * (e.g. inside its own _CItan argument-reduction fallback).
 */

EXPORT __declspec(naked) void _adj_fprem(void)
{
    _asm
    {
	loop_fprem:
		fprem
		fstsw ax
		sahf
		jp loop_fprem
		retn
    }
}

EXPORT __declspec(naked) void _adj_fprem1(void)
{
    _asm
    {
	loop_fprem1:
		fprem1
		fstsw ax
		sahf
		jp loop_fprem1
		retn
    }
}

/**
 * @brief			NOT YET IMPLEMENTED -- confirmed via real disassembly to be
 *					structurally different from the other _adj_fdiv_* functions: a
 *					64-way jump table dispatching on FPU status/condition-code bits,
 *					1183 bytes long, not a simple per-operand-type divide adjustment.
 *					Left as a stub rather than guessed at; see
 *					reference_disasm_vb_tutorial_abi_notes.md for what was confirmed.
 */
EXPORT void __stdcall _adj_fdiv_r(void)
{
    LOG(LOG_DEBUG) << L"_adj_fdiv_r: not implemented";
}

/*
 * The rtcXxx family below is what VB6's Sin/Cos/Tan/Atn/Log/Exp/Sqr language keywords
 * actually compile calls to -- NOT the _CIxxx thunks above directly (those are a
 * lower layer, presumably used by rtcSin/rtcCos/rtcTan internally in the real DLL, and
 * possibly by other directly-typed-Double codegen paths). These were entirely missing
 * from this project (no Source.def entry at all), discovered while testing the
 * _CIxxx fixes above via a real compiled VB6 Sin()/Cos()/... call, which failed to
 * load with "can't find ordinal 582" (rtcSin). Confirmed via real msvbvm60.dll
 * disassembly for all seven. Each takes/returns a plain `double` via the normal
 * stdcall convention (unlike the FPU-stack-only _CIxxx functions), so no naked asm is
 * needed -- a normal C++ function already produces the right convention (MSVC returns
 * `double` via ST(0) for every calling convention, matching what the real functions
 * do too).
 *
 * Confirmed real DLL behavior: on a domain violation, it calls vbaRaiseException and
 * then CONTINUES ON to compute and return a value anyway (there's no early return in
 * the real disassembly) -- presumably relying on the exception mechanism itself
 * (Win32 SEH under the hood) to redirect control flow if there's a handler (e.g. "On
 * Error Resume Next"), and otherwise just falling through. Matched here by calling
 * vbaRaiseException without returning early.
 */

/**
 * @brief			Implements the VB Sin() function.
 */
EXPORT double __stdcall rtcSin(double value)
{
    if (fabs(value) >= 9.223372036854776e18)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    return sin(value);
} /* rtcSin */

/**
 * @brief			Implements the VB Cos() function.
 */
EXPORT double __stdcall rtcCos(double value)
{
    if (fabs(value) >= 9.223372036854776e18)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    return cos(value);
} /* rtcCos */

/**
 * @brief			Implements the VB Tan() function.
 */
EXPORT double __stdcall rtcTan(double value)
{
    if (fabs(value) >= 9.223372036854776e18)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    return tan(value);
} /* rtcTan */

/**
 * @brief			Implements the VB Atn() function. No domain restriction (confirmed
 *					-- the real DLL doesn't check anything for this one).
 */
EXPORT double __stdcall rtcAtn(double value)
{
    return atan(value);
} /* rtcAtn */

/**
 * @brief			Implements the VB Sqr() function. Confirmed: negative input raises
 *					Invalid procedure call.
 */
EXPORT double __stdcall rtcSqr(double value)
{
    if (value < 0.0)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    return sqrt(value);
} /* rtcSqr */

/**
 * @brief			Implements the VB Log() function (natural logarithm). Confirmed:
 *					zero or negative input raises Invalid procedure call.
 */
EXPORT double __stdcall rtcLog(double value)
{
    if (value <= 0.0)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    return log(value);
} /* rtcLog */

/**
 * @brief			Implements the VB Exp() function. Confirmed: input less than -746
 *					raises Invalid procedure call (underflow past what's meaningful);
 *					an overflow from the underlying exp() (errno==ERANGE, i.e. the
 *					input was too large positive) raises Overflow.
 */
EXPORT double __stdcall rtcExp(double value)
{
    if (value < -746.0)
    {
        vbaRaiseException(VBA_EXCEPTION_INVALID_PROCEDURE_CALL);
    }

    errno = 0;
    double result = exp(value);

    if (errno == ERANGE)
    {
        vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
    }

    return result;
} /* rtcExp */
