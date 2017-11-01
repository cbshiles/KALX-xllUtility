// evaluate.cpp - evaluate an expression
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include <regex>
#include "utility.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif 

using namespace std;
using namespace xll;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xstring xstring;

template<class X>
inline XOPER<X>
bangize(const XOPER<X>& oFor)
{
	static xstring fmt(_T("!$1"));
	static basic_regex<xchar> re(_T("(?!!)\\b(\\$?[A-Z]{1,3}\\$?\\d+(:\\$?[A-Z]{1,3}\\$?\\d+)?)\\b"));

	xstring s(oFor.val.str + 1, oFor.val.str[0]);
	s = regex_replace(s, re, fmt);

	LXOPER<X> xName = ExcelX(xlfNames, XMissing<X>(), XOPER<X>(3));
	if (xName.xltype == xltypeMulti) {
		xstring pat(_T("(?!!)\\b("));
		pat.append(xName[0].val.str + 1, xName[0].val.str[0]);
		for (xword i = 1; i < size<X>(xName); ++i) {
			pat.append(_T("|"));
			pat.append(xName[i].val.str + 1, xName[i].val.str[0]);
		}
		pat.append(_T(")\\b"));

		// Match named ranges.
		s = regex_replace(s, basic_regex<xchar>(pat), fmt);
	}

	// remove multiple bangs
	static basic_regex<xchar> bb(_T("!+"));
	static xstring b(_T("!"));
	s = regex_replace(s, bb, b);

	return XOPER<X>(s.c_str(), static_cast<xchar>(s.length()));
}


static AddInX X_(xai_eval)(
	FunctionX(XLL_LPXLOPERX, TX_("?xll_eval"), UTILITY_PREFIX _T("EVALUATE"))
	.Arg(XLL_LPOPERX XLL_UNCALCEDX, _T("Expr"), _T("is an expression to be evaluated. "), _T("1 + 1"))
	.Category(CATEGORY)
	.FunctionHelp(_T("This function calls xlfEvaluate on Expr."))
	.Documentation(
		_T("This has the same effect as pressing F9 on selected text in the forumla bar. Alias <codeInline>EVAL</codeInline>. ")
	)
	.Alias(_T("EVAL"))
);
LPXLOPERX WINAPI
X_(xll_eval)(LPOPERX px)
{
#pragma XLLEXPORT
	static OPERX x;

	x = ExcelX(xlfEvaluate, bangize(*px));

	return &x;
}

#ifdef _DEBUG

int
test_bangize(void)
{
	try {
		ensure (bangize(OPERX("xyz")) == _T("xyz"));
		ensure (bangize(OPERX("xyz")) == _T("xyz"));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenX> xao_test_bangize(test_bangize);


#endif // _DEBUG