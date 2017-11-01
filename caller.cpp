// caller.cpp - who's calling?
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif

using namespace xll;

static AddInX X_(xai_caller)(
	FunctionX(XLL_BOOLX XLL_UNCALCEDX, TX_("?xll_caller"), UTILITY_PREFIX _T("CALLER"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return TRUE if the cell is being called by F2 Enter."))
	.Documentation(
		_T("Useful as a condition to Excel's built-in function IF ")
		_T("to reset a value. Alias <codeInline>CALLING</codeInline>. ")
	)
	.Alias(_T("CALLING"))
);
BOOL WINAPI
X_(xll_caller)(void)
{
#pragma XLLEXPORT
	static OPERX x;

	x = ExcelX(xlCoerce, ExcelX(xlfCaller));

	return x == 0;
}