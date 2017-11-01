// utility.cpp - Various useful Excel routines.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"
//#include "document.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif 

using namespace xll;

#ifdef _DEBUG

static AddInX xai_utility(
	DocumentX(CATEGORY)
	.Documentation(
		_T("Various useful Excel routines. ")
		_T("See <codeInline>MACROFUN.HLP</codeInline> for complete documentation on all the Excel SDK functions and macros. ")
//		,
//		xml::externalLink(_T("MACROFUN.HLP"), _T("http://support.microsoft.com/kb/128185"))
	)
);

//#define XLL_XL_(fn, ...) Excel<XLOPERX>(xl##fn, __VA_ARGS__)

void NewWorksheet(const OPER& oName)
{
	XLL_XLC(WorkbookInsert, OPER(1)); // new worksheet
	XLL_XLC(WorkbookName, XLL_XLF(GetDocument, OPER(1)), oName);

	// first column width 3
	XLL_XLC(ColumnWidth, OPER(3), OPER("C1:C1"));
	// turn off formulas and gridlines
	XLL_XLC(Display, OPER(false), OPER(false));

}

void setup(void)
{
	XLL_XLC(DefineStyle, OPER("Input"), OPER(2), OPER("General")); // number format

	XLL_XLC(DefineStyle, OPER("Output"), OPER(6), OPER(), OPER(), OPER(16)); // patterns format
//	GRAY_25 = 15,

	XLL_XLC(DefineStyle, OPER("Header"), OPER(3), OPER(), OPER(14)); // font format
}

void addin_sheet(void)
{
	NewWorksheet(OPER("AddIn"));
	XLL_MOVE(1, 1);
	XLL_XLC(Formula, OPER("ADDIN.LIST"));
	XLL_XLC(ApplyStyle, OPER("Header"));

	OPER al = XLL_XLF(Evaluate, OPER("=ADDIN.LIST()"));
	OPER ref(3, 1, al.rows(), static_cast<xcol>(al.columns()));
	XLL_XLC(FormulaArray, OPER("=ADDIN.LIST()"), ref);
	XLL_XLC(Select, ref);
	XLL_XLC(ApplyStyle, OPER("Output"));
}

// Macro to create sample spreadsheet.
static AddIn xai_xllutility("?xll_xllutility", "XLL.UTILITY");
int WINAPI
xll_xllutility(void)
{
#pragma XLLEXPORT
	try {
		setup();
		addin_sheet();

		// save as
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

#endif