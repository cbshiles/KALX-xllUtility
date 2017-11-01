// volatile.cpp - make a function volatile
#include "utility.h"

using namespace xll;

static AddInX X_(xai_volatize)(
	FunctionX(XLL_LPOPERX XLL_VOLATILEX, TX_("?xll_volatize"), UTILITY_PREFIX _T("VOLATILE"))
	.Arg(XLL_LPOPERX , _T("Arg"), _T("is any cell or range. "))
	.Category(_T("Utility"))
	.FunctionHelp(_T("Cause Arg to be evaluated on every recalculation."))
	.Documentation(
		_T("Alias <codeInline>VOLATIZE</codeInline>.")
	)
	.Alias(_T("VOLATIZE"))
);
LPXLOPERX WINAPI
X_(xll_volatize)(LPXLOPERX px)
{
#pragma XLLEXPORT

	return px;
}

#ifndef EXCE12
#define EXCEL2
#include "xll/xll.h"
#pragma message("macro X_(foo) = " XLL_STRZ_(X_(foo)))
//## X_(foo))
//#include "volatile.cpp"
#endif
