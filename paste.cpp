// paste.cpp - Paste arguments into Excel given function name or register id.
// If it is a number, just paste the default arguments and do not define names.
// If it is a string, use that as a rdb prefix for defined names.
#include "utility.h"
#include "document.h"

using namespace xll;

typedef traits<XLOPERX>::xword xword;


enum Style {
	STYLE_INPUT,
	STYLE_OUTPUT,
	STYLE_HANDLE,
	STYLE_OPTIONAL
};
OPER StyleName[] = {
	OPER("Input"),
	OPER("Output"),
	OPER("Handle"),
	OPER("Optional")
};
inline void
ApplyStyle(Style s)
{
	Excel<XLOPER>(xlcDefineStyle, StyleName[s]);
	Excel<XLOPER>(xlcApplyStyle, StyleName[s]);
}


/*
static AddIn xai_xll_paste(
	MacroX("_xll_paste@0", "XLL.PASTE.FUNCTION")
	.Category(CATEGORY)
	.FunctionHelp(_T("Paste a function into Excel. "))
	.Documentation(
		"When the active cell is a register id, paste the default arguments below the active cell "
		"and replace the current cell with a call to the function using relative arguments. "
		"When the active cell is a string, look for the register id or a call to the correponding function "
		"in the cell to its right, paste the argument names below the active cell, and the default values "
		"to their right and name them using the active cell contents as a prefix. "
		"Replace the register id/function call with a function call using the named ranges. "
	)
);
extern "C" int __declspec(dllexport) WINAPI
xll_paste(void)
{
	int result(0);

	try {
		Excel<XLOPER>(xlcEcho, OPER(false));

		OPER oAct = Excel<XLOPER>(xlCoerce, Excel<XLOPER>(xlfActiveCell));
		
		if (oAct.xltype == xltypeNum)
			PasteRegidx();
		else if (oAct.xltype == xltypeStr)
			PasteNameX();
		else
			throw std::runtime_error("XLL.PASTE.FUNCTION: Active cell must be a number or a string. ");

		Excel<XLOPER>(xlcEcho, OPER(true));

	}
	catch (const std::exception& ex) {
		Excel<XLOPER>(xlcEcho, OPER(true));
		XLL_ERROR(ex.what());

		return 0;
	}

	return result;
}

int
xll_paste_close(void)
{
	try {
		if (Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Paste Function")))
			Excel<XLOPER>(xlfDeleteCommand, OPER(7), OPER(4), OPER("Paste Function"));
	}
	catch (const std::exception& ex) {
		XLL_INFO(ex.what());
		
		return 0;
	}

	return 1;
}
static Auto<Close> xac_paste(xll_paste_close);

int
xll_paste_open(void)
{
 	try {
		// Try to add just above first menu separator.
		OPER oPos;
		oPos = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("-"));
		oPos = 5;

		OPER oAdj = Excel<XLOPER>(xlfGetBar, OPER(7), OPER(4), OPER("Paste Function"));
		if (oAdj == Err(xlerrNA)) {
			OPER oAdj(1, 5);
			oAdj(0, 0) = "Paste Function";
			oAdj(0, 1) = "XLL.PASTE.FUNCTION";
			oAdj(0, 3) = "Paste function under cursor into spreadsheet.";
			Excel<XLOPER>(xlfAddCommand, OPER(7), OPER(4), oAdj, oPos);
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_paste(xll_paste_open);
*/

inline void
Move(short r, short c)
{
	ExcelX(xlcSelect, ExcelX(xlfOffset, ExcelX(xlfActiveCell), OPERX(r), OPERX(c)));
}

static AddInX xai_paste_basic(
	MacroX(_T("?xll_paste_basic"), _T("XLL.PASTE.BASIC"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Paste a function with default arguments. Shortcut Ctrl-Shift-B."))
	.Documentation(_T("Does not define names. "))
);
int WINAPI
xll_paste_basic(void)
{
#pragma XLLEXPORT
	ExcelX(xlfEcho, OPERX(false));

	try {
		OPER oRegId = Excel<XLOPER>(xlCoerce, Excel<XLOPER>(xlfActiveCell));
		ensure (oRegId.xltype == xltypeNum);

		PasteRegidX(oRegId.val.num);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}
	ExcelX(xlfEcho, OPERX(true));

	return 1;
}
// Ctrl-Shift-B
static On<Key> xok_paste_basic(_T("^+B"), _T("XLL.PASTE.BASIC"));

// create named ranges for arguments
static AddInX xai_paste_create(
	MacroX(_T("?xll_paste_create"), _T("XLL.PASTE.CREATE"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Paste a function and create named ranges for arguments. Shortcut Ctrl-Shift-C"))
	.Documentation(_T("Uses current cell as a prefix. "))
);
int WINAPI
xll_paste_create(void)
{
#pragma XLLEXPORT
	ExcelX(xlfEcho, OPERX(false));

	try {
		OPERX xAct = ExcelX(xlfActiveCell);
		OPERX xPre = ExcelX(xlCoerce, xAct);

		// use cell to right if in cell containing handle
		if (xPre.xltype == xltypeNum) {
			Move(0, -1);
			xAct = ExcelX(xlfActiveCell);
			xPre = ExcelX(xlCoerce, xAct);
		}

//		ensure (xPre.xltype == xltypeStr);
		ExcelX(xlcAlignment, OPERX(4)); // align right

		if (xPre)
			xPre = ExcelX(xlfConcatenate, xPre, OPERX(_T(".")));

		OPERX xFor = ExcelX(xlfGetCell, OPERX(6), ExcelX(xlfOffset, xAct, OPERX(0), OPERX(1))); // formula
		ensure (xFor.xltype == xltypeStr);
		ensure (xFor.val.str[1] == '=');

		// extract "=Function"
		OPERX xFind = ExcelX(xlfFind, OPERX(_T("(")), xFor);
		if (xFind.xltype == xltypeNum)
			xFor = ExcelX(xlfLeft, xFor, OPERX(xFind - 1));

		// get regid
		xFor = ExcelX(xlfEvaluate, xFor);
		ensure (xFor.xltype == xltypeNum);
		double regid = xFor.val.num;

		const ArgsX* pargs = ArgsMapX::Find(regid);

		if (!pargs) {
			XLL_WARNING("XLL.PASTE.CREATE: could not find register id of function");

			return 0;
		}

		xFor = ExcelX(xlfConcatenate, OPERX(_T("=")), pargs->FunctionText(), OPERX(_T("(")));

		for (unsigned short i = 1; i < pargs->ArgCount(); ++i) {
			Move(1, 0);

			ExcelX(xlcFormula, pargs->Arg(i).Name());
			ExcelX(xlcAlignment, OPERX(4)); // align right
			OPERX xNamei = ExcelX(xlfConcatenate, xPre, pargs->Arg(i).Name());

			Move(0, 1);
			ExcelX(xlcDefineName, xNamei);
			
			// paste default argument
			OPERX xDef = pargs->Arg(i).Default();
			if (xDef.xltype == xltypeStr && xDef.val.str[1] == '=') {
				OPERX xEval = ExcelX(xlfEvaluate, xDef);
				if (xEval.size() > 1) {
					OPERX xFor = ExcelX(xlfConcatenate, 
						OPERX(_T("=RANGE.SET(")), 
						OPERX(xDef.val.str + 2, xDef.val.str[0] - 1), 
						OPERX(_T(")")));
					ExcelX(xlcFormula, xFor);
					xNamei = ExcelX(xlfConcatenate, OPERX(_T("RANGE.GET(")), xNamei, OPERX(_T(")")));
				}
				else {
					ExcelX(xlcFormula, xDef);
				}
			}
 			else {
				ExcelX(xlcFormula, xDef);
			}
			// style
			if (pargs->Arg(i).Name().val.str[1] == '_')
				ApplyStyle(STYLE_OPTIONAL);
			else 
				ApplyStyle(STYLE_OUTPUT);

			if (i > 1) {
				xFor = ExcelX(xlfConcatenate, xFor, OPERX(_T(", ")));
			}
			xFor = ExcelX(xlfConcatenate, xFor, xNamei);

			Move(0, -1);
		}
		xFor = ExcelX(xlfConcatenate, xFor, OPERX(_T(")")));

		ExcelX(xlcSelect, xAct);
		Move(0, 1);
		ExcelX(xlcFormula, xFor);
		Move(0, -1);

		ExcelX(xlcSelect, ExcelX(xlfOffset, xAct, OPER(0), OPER(0), OPER(pargs->ArgCount() + 1), OPER(2))); // select range for RDB.DEFINE
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	ExcelX(xlfEcho, OPERX(true));

	return 1;
}
// Ctrl-Shift-C
static On<Key> xok_paste_create(_T("^+C"), _T("XLL.PASTE.CREATE"));
