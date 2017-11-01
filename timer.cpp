// timer.cpp - timer
#include "xll/utility/timer.h"
#include "utility.h"

using namespace xll;

static AddInX X_(xai_timer_start)(
	FunctionX(XLL_HANDLEX, TX_("?xll_timer_start"), _T("TIMER.START"))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Start a timer."))
	.Documentation()
);
HANDLEX WINAPI
X_(xll_timer_start)(void)
{
#pragma XLLEXPORT

	handle<utility::timer> h(new utility::timer());

	h->start();

	return h.get();
}

static AddInX X_(xai_timer_stop)(
	FunctionX(XLL_HANDLEX, TX_("?xll_timer_stop"), _T("TIMER.STOP"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a timer. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Stops a timer."))
	.Documentation()
);
HANDLEX WINAPI
X_(xll_timer_stop)(HANDLEX h)
{
#pragma XLLEXPORT

	try {
		handle<utility::timer> h_(h);
		ensure (h_);

		h_->stop();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return h;
}

static AddInX X_(xai_timer_reset)(
	FunctionX(XLL_HANDLEX, TX_("?xll_timer_reset"), _T("TIMER.RESET"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a timer. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Resets a timer."))
	.Documentation()
);
HANDLEX WINAPI
X_(xll_timer_reset)(HANDLEX h)
{
#pragma XLLEXPORT

	try {
		handle<utility::timer> h_(h);
		ensure (h_);

		h_->reset();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return h;
}

static AddInX X_(xai_timer_elapsed)(
	FunctionX(XLL_DOUBLE, TX_("?xll_timer_elapsed"), _T("TIMER.ELAPSED"))
	.Arg(XLL_HANDLEX, _T("Handle"), _T("is a handle to a timer. "))
	.Volatile()
	.Category(CATEGORY)
	.FunctionHelp(_T("Elapsed time a timer has been running."))
	.Documentation()
);
double WINAPI
X_(xll_timer_elapsed)(HANDLEX h)
{
#pragma XLLEXPORT

	handle<utility::timer> h_(h);

	return h_ ? h_->elapsed() : 0;
}
