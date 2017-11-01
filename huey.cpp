// huey.cpp - Baby Huey bangs on your spreadsheet
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "utility.h"
#include "xll/utility/timer.h"

using namespace xll;
using utility::timer;

static long count(0);
static long modulo(1);
static timer timer_;
static enum huey_states { STOPPED, RUNNING, PAUSED } state(STOPPED);

void check_abort(void)
{
	if (!!XLL_XL_(Abort)) {
		state = STOPPED;

		return;
	}
	double speed = count/timer_.elapsed();
	modulo = static_cast<long>(speed*1.0); // check abort every 1.0 seconds.
	if (modulo == 0)
		modulo = 1;
}

static AddInX xai_baby_huey(
	MacroX(_T("?xll_baby_huey"), _T("HUEY.RUN"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Baby Huey bangs on your workbook."))
	.Documentation(
		_T("Repeatedly calls <codeInline>CALCULATE.NOW</codeInline> (F9) until " )
		_T("the Escape key is pressed or the macro is stopped with either <codeInline>HUEY.STOP</codeInline> ")
		_T("or paused with <codeInline>HUEY.PAUSE</codeInline>. ")
/*		,
		xml::element()
		.content(xml::xlink(_T("HUEY.PAUSE")))
		.content(xml::xlink(_T("HUEY.STOP")))
		.content(xml::externalLink(_T("Baby Huey"), _T("http://en.wikipedia.org/wiki/Baby_Huey")))
*/	)
);
int WINAPI
xll_baby_huey(void)
{
#pragma XLLEXPORT
	static bool shown(false);

	if (!shown)
		Excel<XLOPERX>(xlcAlert, OPERX(_T("Press the Escape key to stop.")));
	shown = true;

	if (state == STOPPED) {
		count = 0;
		timer_.reset();
	}

	state = RUNNING;
	timer_.start();
	XLL_XLC(Echo, OPERX(false));
	while (state == RUNNING) {
		++count;
		Excel<XLOPER>(xlcCalculateNow);
		if (count % modulo == 0) {
			XLL_XLC(Echo, OPERX(true));
			check_abort();
			XLL_XLC(Echo, OPERX(false));
		}
	}
	timer_.stop();
	XLL_XLC(Echo, OPERX(true));
	if (state == RUNNING) // Esc hit
		state = STOPPED;

	return 1;
}

static AddInX xai_huey_elapsed(
	FunctionX(XLL_DOUBLEX XLL_VOLATILEX , _T("?xll_huey_elapsed"), _T("HUEY.ELAPSED"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Elapsed time in seconds since HUEY.RUN."))
	.Documentation(
		_T("The clock stops ticking when <codeInline>HUEY.PAUSE(TRUE)</codeInline> is called. ")
/*		,
		xml::element()
		.content(xml::xlink(_T("HUEY.PAUSE")))
		.content(xml::xlink(_T("HUEY.RUN")))
		.content(xml::xlink(_T("HUEY.STOP")))
*/	)
);
double WINAPI
xll_huey_elapsed(void)
{
#pragma XLLEXPORT

	return timer_.elapsed();
}

static AddInX xai_huey_count(
	FunctionX(XLL_LONGX XLL_VOLATILEX , _T("?xll_huey_count"), _T("HUEY.COUNT"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Number of times HUEY.RUN has recalculated your spreadsheet."))
	.Documentation(
		_T("This number is not reset when <codeInline>HUEY.PAUSE(TRUE)</codeInline> is called. ")
/*		,
		xml::element()
		.content(xml::xlink(_T("HUEY.PAUSE")))
		.content(xml::xlink(_T("HUEY.RUN")))
		.content(xml::xlink(_T("HUEY.STOP")))
*/	)
);
LONG WINAPI
xll_huey_count(void)
{
#pragma XLLEXPORT

	return count;
}

static AddInX xai_huey_stop(
	FunctionX(XLL_BOOLX , _T("?xll_huey_stop"), _T("HUEY.STOP"))
	.Arg(XLL_BOOLX, _T("Condition"), _T("is a boolean value indicating when to stop Huey. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Stop the HUEY.RUN macro."))
	.Documentation(
		_T("This resets the count and elapsed time. ")
/*		,
		xml::element()
		.content(xml::xlink(_T("HUEY.COUNT")))
		.content(xml::xlink(_T("HUEY.ELAPSED")))
		.content(xml::xlink(_T("HUEY.PAUSE")))
		.content(xml::xlink(_T("HUEY.RUN")))
*/	)
);
BOOL WINAPI
xll_huey_stop(BOOL b)
{
#pragma XLLEXPORT

	if (b && state == RUNNING) {
		state = STOPPED;
	}

	return b;
}

static AddInX xai_huey_pause(
	FunctionX(XLL_BOOLX , _T("?xll_huey_pause"), _T("HUEY.PAUSE"))
	.Arg(XLL_BOOLX, _T("Condition"), _T("is a boolean value indicating when to pause Huey. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Pause the HUEY.RUN macro."))
	.Documentation(
		_T("Resume with Alt-F8, HUEY.RUN. ")
/*		,
		xml::element()
		.content(xml::xlink(_T("HUEY.STOP")))
		.content(xml::xlink(_T("HUEY.RUN")))
*/	)
);
BOOL WINAPI
xll_huey_pause(BOOL b)
{
#pragma XLLEXPORT

	if (b && state == RUNNING)
		state = PAUSED;

	return b;
}
