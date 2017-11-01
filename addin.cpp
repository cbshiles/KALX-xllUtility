// addin.cpp - expose AddIn and Args interfaces
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"
//#include "document.h"

#define IS_ADDIN _T("is the Excel function text name of an add-in")

using namespace xll;

static AddInX X_(xai_addin_list)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_list"), _T("ADDIN.LIST"))
	.Arg(XLL_PSTRINGX, _T("_Pattern"), _T("Is an optional wildcard pattern to match. "), _T("*"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Get a list of the Excel FunctionText name for all registered add-ins."))
#ifdef XLLUTILITY
	.Documentation(
		_T("Only the functions in the dll in which this was compiled will show up. ")
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_list)(xcstr pat)
{
#pragma XLLEXPORT
	static OPERX oResult;

	try {
		oResult = OPERX();
		for (AddInX::addin_citer i = AddInX::List().begin(); i != AddInX::List().end(); ++i) {
			if (!(*i)->Args().isDocument()) {
				if (*pat == 0 || ExcelX(xlfSearch, OPERX(pat + 1, pat[0]), (*i)->Args().FunctionText()) == 1)
					oResult.push_back((*i)->Args().FunctionText());
			}
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oResult = ErrX(xlerrValue);
	}

	return &oResult;
}

static AddInX X_(xai_addin_isfunction)(
	FunctionX(XLL_BOOLX, TX_("?xll_addin_isfunction"), _T("ADDIN.ISFUNCTION"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.ISFUNCTION"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a boolean indicating whether or not add-in is a function."))
#ifdef XLLUTILITY
	.Documentation()
#endif
);
BOOL WINAPI
X_(xll_addin_isfunction)(xcstr text)
{
#pragma XLLEXPORT
	BOOL b(FALSE);

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		b = a.isFunction();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return b;
}

static AddInX X_(xai_addin_ismacro)(
	FunctionX(XLL_BOOLX, TX_("?xll_addin_ismacro"), _T("ADDIN.ISMACRO"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.ISMACRO"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a boolean indicating whether or not add-in is a macro."))
#ifdef XLLUTILITY
	.Documentation()
#endif
);
BOOL WINAPI
X_(xll_addin_ismacro)(xcstr text)
{
#pragma XLLEXPORT
	BOOL b(FALSE);

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		b = a.isMacro();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return b;
}

static AddInX X_(xai_addin_argcount)(
	FunctionX(XLL_LONGX, TX_("?xll_addin_argcount"), _T("ADDIN.ARGCOUNT"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.ARGCOUTN"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the number of aguments add-in takes."))
#ifdef XLLUTILITY
	.Documentation(
		_T("The argument count does not include the return value. ")
	)
#endif
);
LONG WINAPI
X_(xll_addin_argcount)(xcstr text)
{
#pragma XLLEXPORT
	int n(-1);

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		n = a.ArgCount();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return n;
}

static AddInX X_(xai_addin_arg)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_arg"), _T("ADDIN.ARG"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.ARG"))
	.Arg(XLL_WORDX, _T("index"), _T("is the 1-based index of the argument "), 1)
	.Category(CATEGORY)
	.FunctionHelp(
		_T("Returns a three column array of the type, name, and help string of the add-in argument ")
		_T("specified by index.")
	)
#ifdef XLLUTILITY
	.Documentation(
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_arg)(xcstr text, xword i)
{
#pragma XLLEXPORT
	static OPERX oArg;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		const ArgX& arg = a.Arg(i);
		oArg.resize(1, 3);
		oArg[0] = ArgType(arg.Type());
		oArg[1] = arg.Name();
		oArg[2] = arg.Help();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oArg = ErrX(xlerrNA);
	}

	return &oArg;
}

static AddInX X_(xai_addin_procedure)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_procedure"), _T("ADDIN.PROCEDURE"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.PROCEDURE"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the C procedure that the add-in calls."))
#ifdef XLLUTILITY
	.Documentation(
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_procedure)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oProc;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oProc = a.Procedure();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oProc = ErrX(xlerrNA);
	}

	return &oProc;
}

static AddInX X_(xai_addin_typetext)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_typetext"), _T("ADDIN.TYPETEXT"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.TYPETEXT"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the type text of the add-in."))
#ifdef XLLUTILITY
	.Documentation(
		_T("The type text encodes the return type of <codeInline>AddIn</codeInline>, ")
		_T("the argument signature, and any function modifiers. ")
		,
		xml::externalLink(_T("Using the <codeInline>CALL</codeInline> and <codeInline>REGISTER</codeInline> functions."),
		_T("http://office.microsoft.com/en-us/excel-help/using-the-call-and-register-functions-HP010062480.aspx"))
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_typetext)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oText;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oText = a.TypeText();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oText = ErrX(xlerrNA);
	}

	return &oText;
}

static AddInX X_(xai_addin_argtext)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_argtext"), _T("ADDIN.ARGTEXT"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.ARGTEXT"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the argument text of the add-in."))
#ifdef XLLUTILITY
	.Documentation(
		_T("The argument text is what Excel displays when you hit Ctrl-Shift-A after typing the ")
		_T("function text of the <codeInline>add-in</codeInline>. ")
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_argtext)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oText;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oText = a.ArgumentText();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oText = ErrX(xlerrNA);
	}

	return &oText;
}

static AddInX X_(xai_addin_category)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_category"), _T("ADDIN.CATEGORY"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.CATEGORY"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the functon category to which the add-in belongs."))
#ifdef XLLUTILITY
	.Documentation(
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_category)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oText;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oText = a.Category();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oText = ErrX(xlerrNA);
	}

	return &oText;
}

static AddInX X_(xai_addin_helptopic)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_helptopic"), _T("ADDIN.HELPTOPIC"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.HELPTOPIC"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the full path and topic id of the help file for the add-in."))
#ifdef XLLUTILITY
	.Documentation(
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_helptopic)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oText;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oText = a.HelpTopic();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oText = ErrX(xlerrNA);
	}

	return &oText;
}

static AddInX X_(xai_addin_functionhelp)(
	FunctionX(XLL_LPOPERX, TX_("?xll_addin_functionhelp"), _T("ADDIN.FUNCTIONHELP"))
	.Arg(XLL_CSTRINGX, _T("AddIn"), IS_ADDIN, _T("ADDIN.FUNCTIONHELP"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Returns a the function help displayed in the Function Wizard for the add-in."))
#ifdef XLLUTILITY
	.Documentation(
	)
#endif
);
LPOPERX WINAPI
X_(xll_addin_functionhelp)(xcstr text)
{
#pragma XLLEXPORT
	static OPERX oText;

	try {
		const ArgsX& a(AddInX::FindFunctionText(text)->Args());
		oText = a.FunctionHelp();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		oText = ErrX(xlerrNA);
	}

	return &oText;
}
