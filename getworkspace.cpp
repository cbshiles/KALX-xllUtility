// getworkspace.cpp - Various methods from GET.WORKSPACE
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

using namespace xll;

static AddInX X_(xai_get_document)(
	FunctionX(XLL_LPOPERX, TX_("?xll_get_document"), _T("GET_DOCUMENT"))
	.Arg(XLL_WORDX, _T("Num"), _T("is a number specifying the type of document information you want."),
		10) // last used row
	.Arg(XLL_PSTRINGX, _T("_Name"), _T("is the name of an open document. "))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns information about the document."))
	.Documentation(
		_T("See <codeInline>GET.DOCUMENT</codeInline> in <codeInline>MACROFUN.HLP</codeInline> for documentation. ")
		,
		xml::externalLink(_T("MACROFUN.HLP"), _T("http://support.microsoft.com/kb/128185"))
	)
);
LPOPERX WINAPI
X_(xll_get_document)(WORD num, xcstr win)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetDocument, OPERX(num), OPERX(win + 1, win[0]));

	return &oResult;
}

static AddInX X_(xai_get_window)(
	FunctionX(XLL_LPOPERX, TX_("?xll_get_window"), _T("GET_WINDOW"))
	.Arg(XLL_WORDX, _T("Num"), _T("is a number specifying the type of window information you want."),
		1) // name of workbook and sheet
	.Arg(XLL_PSTRINGX, _T("_Window"), _T("is the name that appears in the title bar of the window that you want information about. "))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns information about the window."))
	.Documentation(
		_T("See <codeInline>GET.WINDOW</codeInline> in <codeInline>MACROFUN.HLP</codeInline> for documentation. ")
		,
		xml::externalLink(_T("MACROFUN.HLP"), _T("http://support.microsoft.com/kb/128185"))
	)
);
LPOPERX WINAPI
X_(xll_get_window)(WORD num, xcstr win)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetWindow, OPERX(num), OPERX(win + 1, win[0]));

	return &oResult;
}

static AddInX X_(xai_get_workspace)(
	FunctionX(XLL_LPOPERX, TX_("?xll_get_workspace"), _T("GET_WORKSPACE"))
	.Arg(XLL_WORDX, _T("Num"), _T("is a number specifying the type of workspace information you want. "),
		26) // name of user
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns information about the workspace."))
	.Documentation(
		_T("See <codeInline>GET.WORKSPACE</codeInline> in <codeInline>MACROFUN.HLP</codeInline> for documentation. ")
		,
		xml::externalLink(_T("MACROFUN.HLP"), _T("http://support.microsoft.com/kb/128185"))
	)
);
LPOPERX WINAPI
X_(xll_get_workspace)(WORD num)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetWorkspace, OPERX(num));

	return &oResult;
}

static AddInX X_(xai_get_procedures)(
	FunctionX(XLL_LPOPERX XLL_UNCALCEDX, TX_("?xll_get_procedures"), _T("GET.PROCEDURES"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get a 3 column array of all currently registered procedures in dynamic link libraries."))
	.Documentation(
		_T("The first column contains the names of the DLLs that contain the procedures. ")
		_T("The second column contains the names of the procedures in the DLLs. ")
		_T("The third column contains text strings specifying the data types of the return values, and the number and data types of the arguments. ")
		,
		xml::externalLink(_T("Using the <codeInline>CALL</codeInline> and <codeInline>REGISTER</codeInline> functions."),
		_T("http://office.microsoft.com/en-us/excel-help/using-the-call-and-register-functions-HP010062480.aspx"))
	)
);
LPOPERX WINAPI
X_(xll_get_procedures)(void)
{
#pragma XLLEXPORT
	static OPERX oResult;

	oResult = ExcelX(xlfGetWorkspace, OPERX(44));

	return &oResult;
}