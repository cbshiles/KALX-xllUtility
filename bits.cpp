// bits.cpp - bitwise operations
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

using namespace xll;

static AddInX xai_bitand(
	FunctionX(XLL_LONGX, _T("?xll_bitand"), _T("BITAND"))
	.Arg(XLL_LONGX, _T("Int1"), _T("is an integer."))
	.Arg(XLL_LONGX, _T("Int2"), _T("is an integer. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the bitwise AND of Int1 and Int2."))
	.Documentation(_T(""))
);
LONG WINAPI xll_bitand(LONG i1, LONG i2)
{
#pragma XLLEXPORT

	return i1&i2;
}

static AddInX xai_bitor(
	FunctionX(XLL_LONGX, _T("?xll_bitor"), _T("BITOR"))
	.Arg(XLL_LONGX, _T("Int1"), _T("is an integer."))
	.Arg(XLL_LONGX, _T("Int2"), _T("is an integer. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the bitwise OR of Int1 and Int2."))
	.Documentation(_T(""))
);
LONG WINAPI xll_bitor(LONG i1, LONG i2)
{
#pragma XLLEXPORT

	return i1|i2;
}

static AddInX xai_bitxor(
	FunctionX(XLL_LONGX, _T("?xll_bitxor"), _T("BITXOR"))
	.Arg(XLL_LONGX, _T("Int1"), _T("is an integer."))
	.Arg(XLL_LONGX, _T("Int2"), _T("is an integer. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the bitwise XOR of Int1 and Int2."))
	.Documentation(_T(""))
);
LONG WINAPI xll_bitxor(LONG i1, LONG i2)
{
#pragma XLLEXPORT

	return i1^i2;
}