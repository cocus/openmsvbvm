#include "vba_internal.h"

#include "Exceptions.hpp"


#define HIDWORD(d)			(d >> 16)

EXPORT __int64 __stdcall __allmul(
	__int64		i64A,
	__int64		i64B
)
{
	__int64 result;

	if (HIDWORD(i64A) | HIDWORD(i64B))
		result = i64A * i64B;
	else
		result = (unsigned int)i64B * (unsigned __int64)(unsigned int)i64A;
	return result;
}

EXPORT double __stdcall __vbaCyMul(
	double		dblA,
	double		dblB
)
{
	return dblA * dblB / 1000.0; /* Because VB treats the Currencies as 1000 times their value */
}

EXPORT unsigned char __fastcall __vbaUI1I2(
	unsigned __int16 uiIn
)
{
	if (uiIn > 0xFFu)
	{
		vbaRaiseException(VBA_EXCEPTION_OVERFLOW);
	}

	return (unsigned char)uiIn;
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

EXPORT void __stdcall _CIatan(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIcos(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIexp(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIlog(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIsin(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CIsqrt(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _CItan(double value)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m16i(uint16_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m64(double v1, double v2)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall __vbaFpI2(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m32(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_m32i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m16i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m32i(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m32(uint32_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdivr_m64(uint64_t val)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fprem(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fprem1(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}

EXPORT void __stdcall _adj_fdiv_r(void)
{
	// TODO: What to do here?
	DEBUG_DECLARE_WIDE_BUFFER_IF_NEEDED();

	DEBUG_WIDE(
		""
	);
}