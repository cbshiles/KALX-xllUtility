// this.cpp - Return caller's content
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif

using namespace xll;

static AddInX X_(xai_this)(
	FunctionX(XLL_LPXLOPERX, TX_("?xll_this"), /*UTILITY_PREFIX*/ _T("THIS"))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the contents of the calling cell."))
	.Documentation(
		_T("The contents are the last calculated value for the cell.")
	)
);
LPXLOPERX WINAPI
X_(xll_this)(void)
{
#pragma XLLEXPORT
	static OPERX x;

	x = ExcelX(xlCoerce, ExcelX(xlfCaller));

	return &x;
}
