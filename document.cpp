// document.cpp - spreadsheet documentation for add-in functions
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#ifdef EXCEL12
#undef EXCEL12
#endif
#include "xll/xll.h"
//#include "document.h"

using std::string;
using namespace xll;

typedef traits<XLOPERX>::xword xword;
typedef traits<XLOPERX>::xrw xrw;
typedef traits<XLOPERX>::xcol xcol;

#define GRAY_10 LIME // hijack this color for a light gray

// put t in active cell
template<class T>
inline OPER
CellValue(const T& t)
{
	OPER ref(XLL_XLF(ActiveCell));

	XLL_XLC(Select, ref);
	XLL_XLC(Formula, OPER(t));

	return ref;
}
template<>
inline OPER
CellValue<OPER>(const OPER& t)
{
	OPER ref(XLL_XLF(Offset, XLL_XLF(ActiveCell), OPER(0), OPER(0), OPER(t.rows()), OPER(t.columns())));

	if (t)
		XLL_XLC(Formula, t);

	return ref;
}

// define various cell styles
int
xll_define_styles(void)
{
	static bool defined(false);

	if (defined)
		return 1;

	try {
		XLL_XLC(EditColor, OPER(GRAY_10), OPER(220), OPER(220), OPER(220));

		XLL_XLC(DefineStyle, OPER("FunctionName"), OPER(DS_FONT_FORMAT), Missing(), OPER(14), 
			Missing(), Missing(), Missing(), Missing(), OPER(GRAY_10)); 

		XLL_XLC(DefineStyle, OPER("FunctionHelp"), OPER(DS_FONT_FORMAT), Missing(), OPER(12), 
			Missing(), Missing(), Missing(), Missing(), OPER(GRAY_50)); 

		XLL_XLC(DefineStyle, OPER("FunctionText"), OPER(DS_FONT_FORMAT), Missing(), OPER(11), 
			Missing(), Missing(), Missing(), Missing(), OPER(GRAY_80)); 

		XLL_XLC(DefineStyle, OPER("LowerBorder"), OPER(DS_BORDER),
			OPER(0), OPER(0), OPER(0), OPER(BS_THIN),
			OPER(0), OPER(0), OPER(0), OPER(GRAY_10));

		XLL_XLC(DefineStyle, OPER("ArgName"), OPER(false), OPER(true), OPER(true), OPER(true), OPER(false), OPER(false));
		XLL_XLC(DefineStyle, OPER("ArgName"), OPER(DS_FONT_FORMAT), Missing(), OPER(10), 
			OPER(true), Missing(), Missing(), Missing(), OPER(GRAY_80)); 
		XLL_XLC(DefineStyle, OPER("ArgName"), OPER(DS_ALIGNMENT), OPER(HA_RIGHT)); 
		XLL_XLC(DefineStyle, OPER("ArgName"), OPER(DS_BORDER),
			OPER(0), OPER(BS_THIN), OPER(0), OPER(0),
			OPER(0), OPER(GRAY_10), OPER(0), OPER(0));

		XLL_XLC(DefineStyle, OPER("ArgHelp"), OPER(DS_FONT_FORMAT), Missing(), OPER(10), 
			Missing(), Missing(), Missing(), Missing(), OPER(GRAY_80)); 

		XLL_XLC(DefineStyle, OPER("Input"), OPER(DS_FONT_FORMAT), Missing(), OPER(10), 
			Missing(), Missing(), Missing(), Missing(), OPER(GRAY_80)); 
		XLL_XLC(DefineStyle, OPER("Optional"), OPER(false), OPER(false), OPER(false), OPER(false), OPER(false), OPER(false));
		XLL_XLC(DefineStyle, OPER("Output"),   OPER(false), OPER(false), OPER(false), OPER(false), OPER(false), OPER(false));
		XLL_XLC(DefineStyle, OPER("Handle"),   OPER(false), OPER(false), OPER(false), OPER(false), OPER(false), OPER(false));

		defined = true;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
//static Auto<OpenAfter> xao_define_styles(xll_define_styles);

// substring after last period
OPER
ShortName(const OPER& o)
{
	ensure (o.xltype == xltypeStr);

	return o;
	/*

	// skip .get .set
	int n = o.val.str[0];
	bool et = o.val.str[0] > 4
	       && (tolower(o.val.str[n - 2]) == 'e') 
	       && (tolower(o.val.str[n - 1]) == 't');
	int i;
	for (i = 0; i < n; ++i)
		if (o.val.str[n - i] == '.') {
			if (!(i == 3 && et)) 
				break;
		}

	return OPER(o.val.str + n - i + 1, i);
	*/
}

// add a sheet documenting one function
static AddIn xai_doc_one(
	Macro("?xll_doc_one", "DOC.ONE")
	.FunctionHelp("Documents the function in the selected cell.")
);
int WINAPI
xll_doc_one(void)
{
#pragma XLLEXPORT
	OPER o;

	try {
		xll_define_styles();

		o = XLL_XLF(GetCell, OPER(5), XLL_XLF(ActiveCell));
		const Args& po(AddInX::FindFunctionText(o)->Args());
		XLL_XLC(ApplyStyle, OPER("FunctionName"));

		XLL_MOVE(1, 0);
		CellValue(po.FunctionHelp()); 
		XLL_XLC(ApplyStyle, OPER("FunctionHelp"));

		// Macro or Function(Arguments)
		if (po.isFunction()) {
			XLL_MOVE(2,0);
	//		CellValue(po.ReturnType());
			XLL_XLC(ApplyStyle, OPER("LowerBorder"));

			XLL_MOVE(0, 1);
//			xll_paste_regid(po);
			XLL_XLC(ApplyStyle, OPER("Output"));
			if (po.Arg(0).isDate())
				XLL_XLC(FormatNumber, OPER("m/d/yyyy")); //!!! get date format
			else if (po.Arg(0).isHandle())
				XLL_XLC(ApplyStyle, OPER("Handle"));

			XLL_MOVE(0, 1);
			o = po.FunctionText();
			o = XLL_XLF(Concatenate, o, OPER("("), po.ArgumentText(), OPER(")"));
			CellValue(o); 
			XLL_XLC(ApplyStyle, OPER("FunctionText"));

			int max_chars = 0; // guesstimate arg table width
			// args
			XLL_MOVE(1, -2);
			for (xword i = 1; i < po.ArgCount(); ++i) {
				CellValue(po.Arg(i).Name());
				XLL_XLC(ApplyStyle, OPER("ArgName"));

				XLL_MOVE(0, 1);
				OPERX xDef = po.Arg(i).Default();
				if (xDef.xltype == xltypeStr && xDef.val.str[1] == '=') {
					OPERX xEval = ExcelX(xlfEvaluate, xDef);
					if (xEval.size() > 1) {
						OPERX xFor = ExcelX(xlfConcatenate, 
							OPERX(_T("=RANGE.SET(")), 
							OPERX(xDef.val.str + 2, xDef.val.str[0] - 1), 
							OPERX(_T(")")));
						ExcelX(xlcFormula, xFor);
						OPERX xNamei = ExcelX(xlfConcatenate, OPERX(_T("RANGE.GET(")), xNamei, OPERX(_T(")")));
					}
					else {
						ExcelX(xlcFormula, xDef);
					}
				}
				else {
					ExcelX(xlcFormula, xDef);
				}
				CellValue(po.Arg(i).Default());
				if (po.Arg(i).Name().val.str[1] == '_')
					XLL_XLC(ApplyStyle, OPER("Optional"));
				else
					XLL_XLC(ApplyStyle, OPER("Input"));
				if (po.Arg(i).isDate())
					XLL_XLC(FormatNumber, OPER("m/d/yyyy")); // !!!get date format
				else if (po.Arg(i).isHandle())
					XLL_XLC(ApplyStyle, OPER("Handle"));;

				XLL_MOVE(0, 1);
				CellValue(po.Arg(i).Help());
				XLL_XLC(ApplyStyle, OPER("ArgHelp"));

				max_chars = __max(max_chars, XLL_XLF(Len, po.Arg(i).Help()));

				XLL_MOVE(1, -2);
			}
		/*
			// optional description
			OPER oDesc = po.Description();
			if (oDesc.size() > 0) {
				XLL_MOVE(1, 0);
				CellValue(oDesc);
			}
		*/
			XLL_MOVE(-po.ArgCount(), 0);
			OPER ac = XLL_XLF(ActiveCell);
			ac.val.sref.ref.colFirst += 1;
			ac.val.sref.ref.colLast = static_cast<xcol>(ac.val.sref.ref.colFirst + 6);
			XLL_XLC(Select, ac);
			XLL_XLC(Patterns, Num(1), OPER(GRAY_10));
			XLL_XLC(Select, OPER(0,0,1,1));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

// insert sheet for every function in the selection
static AddIn xai_doc_all(
	Macro("?xll_doc_all", "DOC.ALL")
	.FunctionHelp("Add new documented worksheets for all the selected add-in names.")
);
int WINAPI
xll_doc_all(void)
{
#pragma XLLEXPORT
	try {
		OPER o(XLL_XL_(Coerce, XLL_XLF(Selection))); //contents

		ensure (o.xltype == xltypeMulti);

		for (xword i = o.size(); i > 0; --i) {
			XLL_XLC(WorkbookInsert, Num(1));
			XLL_XLC(ColumnWidth, Num(3), OPER("C1:C1"));
//			XLL_XLC(Select, OPER(0, 0, 0x2F, 0x1F)); // all
//			XLL_XLC(Patterns, Num(1), OPER(WHITE));// solid white background
			XLL_XLC(Display, OPER(false), OPER(false));
			XLL_XLC(WorkbookName, XLL_XLF(GetDocument, Num(1)), ShortName(o[i - 1]));
			XLL_XLC(Select, OPER(1, 1, 1, 1)); // B2
			CellValue(o[i-1]);
			ensure (1 == xll_doc_one());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}

static AddIn xai_doc_enum(
	Macro("?xll_doc_enum", "DOC.ENUM")
	.FunctionHelp("Document enumerated values.")
);
int WINAPI
xll_doc_enum(void)
{
#pragma XLLEXPORT
	try {
		OPER o(XLL_XL_(Coerce, XLL_XLF(Selection))); //contents

		ensure (o.xltype == xltypeMulti);

		xll_define_styles();

		XLL_XLC(WorkbookInsert, Num(1));
		XLL_XLC(ColumnWidth, Num(3), OPER("C1:C1"));
		XLL_XLC(Display, OPER(false), OPER(false));
//		XLL_XLC(WorkbookName, XLL_XLF(GetDocument, Num(1)), ShortName(o[i - 1]));
		XLL_XLC(Select, OPER(1, 1, 1, 1)); // B2
//		CellValue(o[i-1]);
		ensure (1 == xll_doc_one());
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
// write C++ code for an add-in