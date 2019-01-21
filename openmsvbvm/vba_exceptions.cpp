
#include "vba_exception.h"

CEXTERN unsigned int __stdcall vbaErrorFromSTGResult(int a1)
{
	unsigned int result; // eax

	if (a1 >= 0)
		return 0;
	if ((a1 & 0x1FFF0000) == 0xA0000)
		return (unsigned __int16)a1;
	if ((a1 & 0x1FFF0000) == 0x10000)
		return 440;
	if (a1 == 0x80004001)
		return 445;
	if (a1 == 0x80030002)
		return 432;
	result = a1;//vbaErrorFromHRESULT(a1);
	if (result > 0x2EA)
		return 440;
	return result;
}


CEXTERN void vbaRaiseException(int exceptionCode)
{
	char debug[300];
	snprintf(debug, sizeof(debug), "%s: exception code = %.4d", __func__, exceptionCode);
	OutputDebugStringA(debug);
}