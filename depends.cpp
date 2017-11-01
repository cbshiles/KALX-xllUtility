// depends.cpp - specify Excel calculation order
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif 

using namespace xll;

static AddInX X_(xai_depends)(
	FunctionX(XLL_LPXLOPERX, TX_("?xll_depends"), UTILITY_PREFIX _T("DEPENDS"))
	.Arg(XLL_LPXLOPERX, _T("Value"), _T("is the value to be returned."))
	.Arg(XLL_LPXLOPERX, _T("Dependent"), _T("is a reference to a cell that is required to be computed before Value."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return Value after Dependent has been computed."))
	.Documentation(
		_T("Excel calculation order is unspecified for independent formulas. This function ")
	    _T("can be used to make calculation order deterministic. The value of ")
		_T("<codeInline>Dependent</codeInline> is not used, only that it has been calculated prior to ")
		_T("returning <codeInline>Value</codeInline>. Alias <codeInline>DEPENDS</codeInline>. ")
	)
	.Alias("DEPENDS")
);
LPXLOPERX WINAPI
X_(xll_depends)(LPXLOPERX pRef, LPXLOPERX pDep)
{
#pragma XLLEXPORT

	return xlretUncalced == traits<XLOPERX>::Excel(xlCoerce, 0, 1, pDep) ? 0 : pRef;
}
