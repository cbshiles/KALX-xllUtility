// document.h - Macros to help create Excel files
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xll/xll.h"

using namespace xll;

#ifdef EX
#undef EX
#endif
#define EX Excel<XLOPER>

// Alignment
enum HAlign { 
	HA_NONE,
	HA_GENERAL,
	HA_LEFT,
	HA_CENTER,
	HA_RIGHT,
	HA_FILL,
	HA_JUSTIFY,
	HA_CENTER_ACROSS_SELECTION
};
enum VAlign {
	VA_NONE,
	VA_TOP,
	VA_CENTER, 
	VA_BOTTOM,
	VA_JUSTIFY
};
enum Orientation {
	O_HORIZONTAL,
	O_VERTICAL,
	O_UPWARD,
	O_DOWNWARD,
	O_AUTOMATIC
};

enum Borders {
	B_OUTLINE, 
	B_LEFT, 
	B_RIGHT, 
	B_TOP, 
	B_BOTTOM
};
enum BorderStyle {
	BS_NONE,
	BS_THIN,
	BS_MEDIUM,
	BS_DASHED,
	BS_DOTTED,
	BS_THICK,
	BS_DOUBLE,
	BS_HAIRLINE
};

// Font
enum FontColor {
	AUTOMATIC, 
	BLACK = 1,
	BROWN = 53,
	OLIVE_GREEN = 52,
	DARK_GREEN = 51,
	DARK_TEAL = 49,
	DARK_BLUE = 11,
	INDIGO = 55,
	GRAY_80 = 56,
	DARK_RED = 9,
	ORANGE = 46,
	DARK_YELLOW = 12,
	GREEN = 10,
	TEAL = 14,
	BLUE = 5,
	BLUE_GRAY = 47,
	GRAY_50 = 16,
	RED = 3,
	LIGHT_ORANGE = 45,
	LIME = 43,
	SEA_GREEN = 50,
	AQUA = 42,
	LIGHT_BLUE = 41,
	VIOLET = 13,
	GRAY_40 = 48,
	PINK = 7,
	GOLD = 44,
	YELLOW = 6,
	BRIGHT_GREEN = 4,
	TURQUOISE = 8,
	SKY_BLUE = 33,
	PLUM = 54,
	GRAY_25 = 15,
	ROSE = 38,
	TAN = 40,
	LIGHT_YELLOW = 36,
	LIGHT_GREEN = 35,
	LIGHT_TURQUOISE = 34,
	PALE_BLUE = 37,
	LAVENDER = 39,
	WHITE = 2,
};

enum FontStyle {
	FS_NONE = 0,
	FS_BOLD = 1, 
	FS_ITALIC = 2, 
	FS_UNDERLINE = 4,
	FS_STRIKE = 8
};

enum DefineStyle {
	DS_NUMBER_FORMAT = 2,
	DS_FONT_FORMAT = 3,
	DS_ALIGNMENT = 4,
	DS_BORDER = 5,
	DS_PATTERN = 6,
	DS_CELL_PROTECTION = 7
};

// helper functions
template<class X>
struct select {
	static XOPER<X> up;
	static XOPER<X> down;
	static XOPER<X> right;
	static XOPER<X> left;
	static XOPER<X> alert;
	static XOPER<X> input;
	static XOPER<X> r_;
	static XOPER<X> c0;
	static XOPER<X> range_set;
	static XOPER<X> range_get;
};

OPER select<XLOPER>::up = OPER("R[-1]C[0]");
OPER select<XLOPER>::down = OPER("R[1]C[0]");
OPER select<XLOPER>::right = OPER("R[0]C[1]");
OPER select<XLOPER>::left = OPER("R[0]C[-1]");
OPER select<XLOPER>::alert = OPER("Would you like to replace the existing definition of ");
OPER select<XLOPER>::input = OPER("Enter the range name.");
OPER select<XLOPER>::r_ = OPER("R[");
OPER select<XLOPER>::c0 = OPER("]C[0]");
OPER select<XLOPER>::range_set = OPER("=RANGE.SET(");
OPER select<XLOPER>::range_get = OPER("=RANGE.GET(");

OPER12 select<XLOPER12>::up = OPER12(L"R[-1]C[0]");
OPER12 select<XLOPER12>::down = OPER12(L"R[1]C[0]");
OPER12 select<XLOPER12>::right = OPER12(L"R[0]C[1]");
OPER12 select<XLOPER12>::left = OPER12(L"R[0]C[-1]");
OPER12 select<XLOPER12>::alert = OPER12(L"Would you like to replace the existing definition of ");
OPER12 select<XLOPER12>::input = OPER12(L"Enter the range name.");
OPER12 select<XLOPER12>::r_ = OPER12(L"R[");
OPER12 select<XLOPER12>::c0 = OPER12(L"]C[0]");
OPER12 select<XLOPER12>::range_set = OPER12(L"=RANGE.SET(");
OPER12 select<XLOPER12>::range_get = OPER12(L"=RANGE.GET(");


#define UP Excel<X>(xlcSelect, XOPER<X>(select<X>::up))
#define DOWN Excel<X>(xlcSelect, XOPER<X>(select<X>::down))
#define RIGHT Excel<X>(xlcSelect, XOPER<X>(select<X>::right))
#define LEFT Excel<X>(xlcSelect, XOPER<X>(select<X>::left))

template<class X>
inline XOPER<X> char1(typename traits<X>::xchar c)
{
	return XOPER<X>(&c, 1);
}
template<class X>
inline XOPER<X>& append(XOPER<X>& x, const XOPER<X> y)
{
	return x = Excel<X>(xlfConcatenate, x, y);
}

// not very efficent lookup
inline const ArgsX*
FindFunctionText(const XLOPERX& text)
{
	for (AddInX::addin_citer i = AddInX::List().begin(); i != AddInX::List().end(); ++i) {
		if (!(*i)->Args().isDocument()) {
			if ((*i)->Args().FunctionText() == text)
				return &((*i)->Args());
		}
	}

	return 0;
}
inline const ArgsX*
FindFunctionText(traits<XLOPERX>::xcstr text)
{
	return FindFunctionText(OPERX(text));
}

inline void
Align(HAlign ha, VAlign va = VA_NONE, bool wrap = false)
{
	if (va == VA_NONE)
		EX(xlcAlignment, OPER(ha), OPER(wrap));
	else
		EX(xlcAlignment, OPER(ha), OPER(wrap), OPER(va));
}

// Range based on active cell.
inline OPER
Range(LONG r, LONG c)
{
	return EX(xlfOffset, EX(xlfActiveCell), OPER(0), OPER(0), OPER(r), OPER(c));
}
// relative move from active cell
inline void
Move(int r, int c)
{
	EX(xlcSelect, EX(xlfOffset, EX(xlfActiveCell), OPER(r), OPER(c)));
}

// put t in active cell
template<class T>
inline OPER
CellValue(const T& t)
{
	OPER ref(EX(xlfActiveCell));

	EX(xlcSelect, ref);
	EX(xlcFormula, OPER(t));

	return ref;
}
template<>
inline OPER
CellValue<OPER>(const OPER& t)
{
	OPER ref(EX(xlfOffset, EX(xlfActiveCell), OPER(0), OPER(0), OPER(t.rows()), OPER(t.columns())));

	if (t)
		EX(xlcFormula, t);

	return ref;
}

#define BOOL(a,b) OPER(((a)&(b))!=0)
inline void
Format(const OPER& font, unsigned int size, FontStyle biu, FontColor c)
{
	EX(xlcFormatFont, font, OPER(size), BOOL(biu, FS_BOLD), BOOL(biu, FS_ITALIC), BOOL(biu, FS_UNDERLINE), BOOL(biu, FS_STRIKE), OPER(c));
}
#undef BOOL

inline void
Border(Borders b, BorderStyle bs, FontColor c = AUTOMATIC)
{
	switch (b) {
	case B_OUTLINE:
		EX(xlcBorder, OPER(bs), Missing(), Missing(), Missing(), Missing(), Missing(), OPER(c));
		break;
	case B_LEFT:
		EX(xlcBorder, Missing(), OPER(bs), Missing(), Missing(), Missing(),  Missing(), Missing(), OPER(c));
		break;
	case B_RIGHT:
		EX(xlcBorder, Missing(), Missing(), OPER(bs), Missing(), Missing(),  Missing(), Missing(), Missing(), OPER(c));
		break;
	case B_TOP:
		EX(xlcBorder, Missing(), Missing(), Missing(), OPER(bs), Missing(),  Missing(), Missing(), Missing(), Missing(), OPER(c));
		break;
	case B_BOTTOM:
		EX(xlcBorder, Missing(), Missing(), Missing(), Missing(), OPER(bs),  Missing(), Missing(), Missing(), Missing(), Missing(), OPER(c));
		break;
	}
}

inline void
NewWorkbook(void)
{
	EX(xlcWorkbookInsert, OPER(1));
	EX(xlcColumnWidth, OPER(3), OPER("C1:C1"));
	// turn off formulas and gridlines
	EX(xlcDisplay, OPER(false), OPER(false));
}

inline void
WorkbookName(const OPER& o)
{
	EX(xlcWorkbookName, EX(xlfGetDocument, OPER(1)), o);

}

// paste argument into xActi, return reference for formula
template<class X> inline 
XOPER<X> PasteDefault(const XOPER<X>& xAct, const XOPER<X>& xActi, const XOPER<X>& xDef)
{
	XOPER<X> xRel = Excel<X>(xlfRelref, xActi, xAct);

	if (xDef && xDef.xltype == xltypeStr && xDef.val.str[0] > 0 && xDef.val.str[1] == '=') {
		XOPER<X> xEval = Excel<X>(xlfEvaluate, xDef);
		if (xEval.size() > 1) {
			XOPER<X> xOff = Excel<X>(xlfOffset, xActi, XOPER<X>(0), XOPER<X>(0), XOPER<X>(1), XOPER<X>(xEval.size()));
			Excel<X>(xlSet, xOff, xEval);

			xRel = Excel<X>(xlfRelref, xOff, xAct);
		}
		else {
			Excel<X>(xlcFormula, xDef);
		}
	}
	else {
		Excel<X>(xlSet, xActi, xDef);
	}

	return xRel;
}

template<class X> inline
void PasteRegid(const XArgs<X>* pargs)
{
	XOPER<X> xAct = Excel<X>(xlfActiveCell);

	XOPER<X> xFor(char1<X>('='));
	append(xFor, pargs->FunctionText());
	append(xFor, char1<X>('('));

	for (unsigned short i = 1; i < pargs->ArgCount(); ++i) {
		DOWN;

		XOPER<X> xActi = Excel<X>(xlfActiveCell);
		XOPER<X> xDef = pargs->Arg(i).Default();
		XOPER<X> xRel = PasteDefault(xAct, xActi, xDef);

		if (i > 1) {
			append(xFor, char1<X>(','));
			append(xFor, char1<X>(' '));
		}
		append(xFor, xRel);
	}
	append(xFor, char1<X>(')'));

	Excel<X>(xlcSelect, xAct);
	Excel<X>(xlcFormula, xFor);
}

void inline
PasteRegidX(double regid)
{
	const Args* parg;
	const Args12* parg12;

	if (0 != (parg = ArgsMap::Find(regid))) {
		if (parg->isFunction())
			return PasteRegid(parg);
	}
	else if (0 != (parg12 = ArgsMap12::Find(regid))) {
		if (parg12->isFunction())
			return PasteRegid(parg12);
	}
	else {
		throw std::runtime_error("XLL.PASTE.FUNCTION: register id not found");
	}
}

template<class X> inline
void PasteName(const XArgs<X>* pargs)
{
	XOPER<X> xAct = Excel<X>(xlfActiveCell);
	XOPER<X> xPre = Excel<X>(xlCoerce, xAct);

	XOPER<X> xFor(char1<X>('='));
	append(xFor, pargs->FunctionText());
	append(xFor, char1<X>('('));

	for (unsigned short i = 1; i < pargs->ArgCount(); ++i) {
		DOWN;

		Excel<X>(xlcFormula, pargs->Arg(i).Name());
		XOPER<X> xNamei = Excel<X>(xlfConcatenate, xPre, pargs->Arg(i).Name());

		RIGHT;
		XOPER<X> xActi = Excel<X>(xlfActiveCell);
		XOPER<X> xDef = pargs->Arg(i).Default();
		XOPER<X> xRel = PasteDefault(xAct, xActi, xDef);
		Excel<X>(xlcDefineName, xNamei, Excel<X>(xlfAbsref, xRel, xAct));
		LEFT;

		if (i > 1) {
			append(xFor, char1<X>(','));
			append(xFor, char1<X>(' '));
		}
		append(xFor, xNamei);
	}
	append(xFor, char1<X>(')'));

	Excel<X>(xlcSelect, xAct);
	RIGHT;
	Excel<X>(xlcFormula, xFor);
	LEFT;
}

void inline
PasteNameX(void)
{
	const Args* parg;
	const Args12* parg12;

	double regid = Excel<XLOPER>(xlCoerce, Excel<XLOPER>(xlfOffset, Excel<XLOPER>(xlfActiveCell), OPER(0), OPER(1))).val.num;

	if (0 != (parg = ArgsMap::Find(regid))) {
		if (parg->isFunction())
			return PasteName(parg);
	}
	else if (0 != (parg12 = ArgsMap12::Find(regid))) {
		if (parg12->isFunction())
			return PasteName(parg12);
	}
	else {
		throw std::runtime_error("XLL.PASTE.FUNCTION: register id not found");
	}
}

/*
inline OPER
ArgValue(const std::string& type, const OPER& o)
{
	if (type == "Bool") {
		ensure (o.xltype == xltypeBool);
	}
	else if (type == "OPER") {
		ensure (o.xltype == xltypeMulti);
	}

	return CellValue(o);
}
*/
/*
// Create and check example add-in call
struct Example {
	Example(const OPER& ai, int n, OPER o[], WORD r = 1, WORD c = 1)
	{
		const Args* pai = FindFunctionText(OPER(ai));
		ensure (pai != 0);
		ensure (n == 1 + pai->ArgCount());

		OPER call(EX(xlfConcatenate, OPER("="), pai->FunctionText()));
		for (int i = 1; i < n; ++i) {
			
			if (i == 1)
				call = EX(xlfConcatenate, call, OPER("("));
			else
				call = EX(xlfConcatenate, call, OPER(", "));

			CellValue(pai->Argument(i));
			Format(Missing(), OPER(12), FS_BOLD, BLACK);
			Align(HA_RIGHT);
			
			Move(0, 1);

			OPER ref = ArgValue(pai->ArgumentType(i), o[i]);
			Format(Missing(), OPER(12), FS_NONE, BLACK);
			EX(xlcDefineName, pai->Argument(i), ref);

			call = EX(xlfConcatenate, call, pai->Argument(i));
			
			Move(ref.rows(), -1);
		}
		call = EX(xlfConcatenate, call, OPER(")"));
		CellValue(pai->FunctionText());
		Format(Missing(), OPER(12), FS_ITALIC, BLACK);
		Align(HA_RIGHT);

		Move(0, 1);
		if (r*c > 1)
			EX(xlcFormulaArray, call, Range(r, c));
		else
			EX(xlcFormula, call);

		// Check value
		if (o[0].xltype != xltypeMissing)
			ensure (o[0] == EX(xlfEvaluate, call));
	}
};
*/
