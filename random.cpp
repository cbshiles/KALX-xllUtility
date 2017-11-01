// random.cpp - generate random OPER's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include <ctime>
#include "utility.h"
#include "xll/utility/registry.h"
#include "xll/utility/srng.h"


#ifndef CATEGORY
#define CATEGORY _T("Test")
#endif

using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xstring xstring;

#define SUBKEY _T("Software\\KALX\\xll")

static utility::srng rng(false);

XLL_ENUM_DOC(xltypeNum, TYPE_NUM, CATEGORY, _T("A 64-bit IEEE floating point number "), _T(""));
XLL_ENUM_DOC(xltypeStr, TYPE_STR, CATEGORY, _T("A character string "), _T(""));
XLL_ENUM_DOC(xltypeBool, TYPE_BOOL, CATEGORY, _T("A Boolean value "), _T(""));
XLL_ENUM_DOC(xltypeRef, TYPE_REF, CATEGORY, _T("A reference to multiple ranges "), _T(""));
XLL_ENUM_DOC(xltypeErr, TYPE_ERR, CATEGORY, _T("An error type "), _T(""));
XLL_ENUM_DOC(xltypeMulti, TYPE_MULTI, CATEGORY, _T("A two dimensional Range of cells "), _T(""));
XLL_ENUM_DOC(xltypeMissing, TYPE_MISSING, CATEGORY, _T("A missing type "), _T(""));
XLL_ENUM_DOC(xltypeNil, TYPE_NIL, CATEGORY, _T("A nil type "), _T(""));
XLL_ENUM_DOC(xltypeSRef, TYPE_SREF, CATEGORY, _T("A reference to a single range "), _T(""));
XLL_ENUM_DOC(xltypeInt, TYPE_INT, CATEGORY, _T("A 16-bit signed integer "), _T(""));

static AddInX X_(xai_rand_num)(
	FunctionX(XLL_DOUBLEX, TX_("?xll_rand_num"), _T("RAND.NUM"))
	.Arg(XLL_DOUBLEX, _T("_Max"), _T("is the optional maximum value."))
	.Arg(XLL_DOUBLEX, _T("_Min"), _T("is the optional minimum value. "))
	.Volatile()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random double."))
	.Documentation(
		_T("This uses the slash (reciprocal of uniform (0, 1]) distribution ")
		_T("and is negative with probability 0.5. ")
	)
);
double WINAPI
X_(xll_rand_num)(double max, double min)
{
#pragma XLLEXPORT
	max = max ? max : 1/rng.real();
	min = min ? min : -1/rng.real();

	return min + (max - min)*rng.real();
}

static AddInX X_(xai_rand_str)(
	FunctionX(XLL_LPOPERX, TX_("?xll_rand_str"), _T("RAND.STR"))
	.Arg(XLL_WORDX, _T("_Length"), _T("is the maximum length of the string. "))
	.Volatile()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random string."))
	.Documentation(
		_T("Maximum string length is 255 (0xFF) in versions prior to Excel 2007 ")
		_T("and 32767 (0x7FFF) after. ")
	)
);
LPXLOPERX WINAPI
X_(xll_rand_str)(xword max)
{
#pragma XLLEXPORT
	static XLOPERX s;
	static xchar t[limits<XLOPERX>::maxchars + 1];

	max = max ? max : -1 + limits<XLOPERX>::maxchars;
	size_t len = rng.between(1, max);
	
	t[0] = (xchar)len;
	for (size_t i = 1; i <= len; ++i) {
		t[i] = static_cast<xchar>(rng.between(32, 127));
	}

	s.xltype = xltypeStr;
	s.val.str = t;

	return &s;
}

static AddInX X_(xai_rand_bool)(
	FunctionX(XLL_BOOLX XLL_VOLATILEX, TX_("?xll_rand_bool"), _T("RAND.BOOL"))
	.Arg(XLL_DOUBLEX, _T("p"), _T("is the probability of returning the value TRUE "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random boolean that is true with probability p."))
	.Documentation(
		_T("If the probability is missing or zero then equal probabilities are used. ")
	)
);
BOOL WINAPI
X_(xll_rand_bool)(double p)
{
#pragma XLLEXPORT
	if (p == 0)
		p = 0.5;

	return rng.real() < p ? FALSE : TRUE;
}

static int err[] = {
    xlerrNull,
    xlerrDiv0,
    xlerrValue,
    xlerrRef,
    xlerrName,
    xlerrNum,
    xlerrNA,
    xlerrGettingData
};

static AddInX X_(xai_rand_err)(
	FunctionX(XLL_LPOPERX XLL_VOLATILEX, TX_("?xll_rand_err"), _T("RAND.ERR"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random error."))
	.Documentation(
		_T("Possible values are #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, and #N/A. ")
	)
);
LPOPERX WINAPI
X_(xll_rand_err)(void)
{
#pragma XLLEXPORT
	static OPERX o;

	o = ErrX(static_cast<WORD>(err[rng.between(0, dimof(err) - 1)]));

	return &o;
}

static xword xltype[] = {
	xltypeNum,
	xltypeBool,
	xltypeErr,
	xltypeInt,
	xltypeStr,
	xltypeMulti
};

LPOPERX WINAPI xll_rand_type(xword type);

static AddInX X_(xai_rand_multi)(
	FunctionX(XLL_LPOPERX XLL_VOLATILEX, TX_("?xll_rand_multi"), _T("RAND.MULTI"))
	.Arg(XLL_WORDX, _T("_Rows"), _T("is the number of rows to return."))
	.Arg(XLL_WORDX, _T("_Columns"), _T("is the number of columns to return. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random two dimensional range."))
	.Documentation(_T("If Rows or Columns is 0, a random size is generated. "))
);
LPOPERX WINAPI
X_(xll_rand_multi)(xword r, xword c) // rows and columns?
{
#pragma XLLEXPORT
	static OPERX o;

	r = r ? r : static_cast<xword>(1/rng.real());
	r = min(r, limits<XLOPERX>::maxrows);
	c = c ? c : static_cast<xword>(1/rng.real());
	c = min(c, limits<XLOPERX>::maxcols);
	if (r == 0)
		r = 1;
	if (c == 0)
		c = 1;
	o.resize(r, c);
	for (xword i = 0; i < o.size(); ++i) {
		xword t = xltype[rng.between(0, dimof(xltype) - 2)]; // not multi
		o[i] = *xll_rand_type(t);
	}

	return &o;
}

static AddInX X_(xai_rand_int)(
	FunctionX(XLL_LPOPERX XLL_VOLATILEX, TX_("?xll_rand_int"), _T("RAND.INT"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random integer."))
	.Documentation(
		_T("This uses the same distribution as <codeInline>RAND.NUM</codeInline> truncated to an integer value. ")
	)
);
LPOPERX WINAPI
X_(xll_rand_int)()
{
#pragma XLLEXPORT
	static OPERX o;

	short int w = static_cast<short int>(1/rng.real());
	if (rng.real() < 0.5)
		w = -w;

	o = IntX(w);

	return &o;
}

static AddInX X_(xai_rand_type)(
	FunctionX(XLL_LPOPERX XLL_VOLATILEX, TX_("?xll_rand_type"), _T("RAND.TYPE"))
	.Arg(XLL_SHORTX, _T("Type"), _T("is the type of OPER to return "), _T("=TYPE_STR()"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a random OPER of the given Type."))
	.Documentation(
		_T("Use values from the TYPE_* enumeration for Type. ")
	)
);
LPOPERX WINAPI
X_(xll_rand_type)(xword type)
{
#pragma XLLEXPORT
	static OPERX o;

	switch (type) {
	case xltypeNum:
		o = xll_rand_num(0,0);

		break;
	case xltypeStr: {
		o = *xll_rand_str(0);

		break;
	}
	case xltypeBool:
		o = (TRUE == xll_rand_bool(0.5));

		break;
	case xltypeErr:
		o = *xll_rand_err();

		break;
	case xltypeInt:
		o = *xll_rand_int();

		break;
	case xltypeMulti:
		o = *xll_rand_multi(0,0);

		break;
	default:
		o = ErrX(xlerrValue);
	}

	return &o;
}

#ifdef _DEBUG

bool
check_strict_weak(const OPERX& x, const OPERX& y, const OPERX& z)
{
	ensure (!(x < x));
	ensure (!(y < y));
	ensure (!(z < z));

	if (x != y && x < y)
		ensure (!(y < x));

	if (y != z && y < z)
		ensure (!(z < y));

	if (x < y && y < z)
		ensure (x < z);

	if (x < y)
		ensure (x < z || z < y);

	return true;
}

int
test_strict_weak()
{
	try {
		OPERX m = *xll_rand_multi(0,0);
		for (xword i = 0; i + 2 < m.size(); ++i)
			check_strict_weak(m[i], m[i + 1], m[i + 2]);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

int
test_rand(void)
{
//	_CrtSetBreakAlloc(3272);
	test_strict_weak();

	return 1;
}
static Auto<OpenAfter> xao_test_rand(test_rand);

#endif // _DEBUG
