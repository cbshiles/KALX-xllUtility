// err.cpp - Excel error types
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"

#ifndef CATEGORY
#define CATEGORY _T("Utility")
#endif

using namespace xll;

XLL_ENUM_DOC(xlerrNull, ERR_NULL, CATEGORY, 
	_T("Intersection of two areas is empty. "), 
	_T("Indicated in Excel by <codeInline>#NULL!</codeInline>. "));
XLL_ENUM_DOC(xlerrDiv0, ERR_DIV0, CATEGORY, 
	_T("Division by zero. "), 
	_T("Indicated in Excel by <codeInline>#DIV/0!</codeInline>. "));
XLL_ENUM_DOC(xlerrValue, ERR_VALUE, CATEGORY, 
	_T("Wrong type of argument or operand. "), 
	_T("Indicated in Excel by <codeInline>#VALUE!</codeInline>. "));
XLL_ENUM_DOC(xlerrRef, ERR_REF, CATEGORY, 
	_T("Cell reference is not valid. "), 
	_T("Indicated in Excel by <codeInline>#REF!</codeInline>. "));
XLL_ENUM_DOC(xlerrName, ERR_NAME, CATEGORY, 
	_T("Unrecognized text in a formula. "), 
	_T("Indicated in Excel by <codeInline>#NAME?</codeInline>. "));
XLL_ENUM_DOC(xlerrNum, ERR_NUM, CATEGORY, 
	_T("Invalid numeric value. "), 
	_T("Indicated in Excel by <codeInline>#NUM!</codeInline>. "));
XLL_ENUM_DOC(xlerrNA, ERR_NA, CATEGORY, 
	_T("A value is not available to a function or a formula. "), 
	_T("Indicated in Excel by <codeInline>#N/A</codeInline>. "));
