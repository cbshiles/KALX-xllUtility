// iferror.cpp - Excel 2003 (and earlier) replacement of 2007 and 2010 IFERROR.
// Copyright (c) 2010 KALX, LLC. All rights reserved. No warranty is made.
#include "xll/xll.h"

using namespace xll;

#ifndef EXCEL12

static AddIn xai_logical(
	Args("Logical")
	.Documentation(_T("Routines related to logic. "))
);

static AddIn xai_iferror(
	Function(XLL_LPOPER, "?xll_iferror", "IFERROR")
	.Arg(XLL_LPOPER, "Value", "is any value or expression or reference.")
	.Arg(XLL_LPOPER, "Value_if_error", "is any value or expression or reference.")
	.Arg(XLL_BOOL, "Empty_is_error", "is an optional boolean indicating empty Values should return Value_if_error. ")
	.Category("Logical")
	.FunctionHelp("Returns value_if_error if expression is an error "
		"and the value of the expression itself otherwise.")
	.Documentation(
		"This is a drop-in replacement for the Excel 2007 and later function of the same name."
	)
);
LPOPER WINAPI
xll_iferror(LPOPER pval, LPOPER perr, BOOL nil)
{
#pragma XLLEXPORT
	static OPER oResult;

	oResult = (pval->xltype == xltypeErr || nil && pval->xltype == xltypeNil) ? *perr : *pval;

	return &oResult;
}

#endif // EXCEL12