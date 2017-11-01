Project Description

Excel add-in containing useful general purpose routines.

Download xllutility.zip, extract all files to any folder, and open the add-in and spreadsheet.
Select a cell containing a function, start the Function Wizard, and click on Help on this function to see the documentation.

You may need to right click on xllutility.chm, choose Properties, and Unblock this file.

Overview

This project exposes AddIn objects to Excel. The ADDIN.* functions allow you to access all the information available such as the list of all add-ins matching a pattern, whether or not it is a function or macro, argument and return type information, and category and function help. This is used by the macros DOC.ONE and DOC.ALL to automate the creation of spreadsheets.

The HUEY.* routines are general purpose routines for testing and measuring the performance of spreadsheets. The macro HUEY.RUN just pounds on F9 in a tight loop until the escape key is pressed or HUEY.STOP(Condition) or HUEY.PAUSE(Condition) is called from a spreadsheet with Condition true. The functions HUEY.COUNT and HUEY.ELAPSED allow you to measure the performance of the spreadsheet. Stopping a simulation will reset the count to 1 and elapsed time to 0 for the next simulation. Pausing a simulation does not affect the count, and pauses the internal timer.

Fine grain timing of individual cells is possible using the TIMER.* functions, with a dose of DEPENDS sprinkled in. The function TIMER.START() returns a handle to a timer. Passing the handle to TIMER.STOP will stop the timer and return the same handle. Use TIMER.ELAPSED to find out how long the timer was running. For example, if cell A1 contains a call to a function taking an argument from cell A2 you can call DEPENDS(TIMER.START(), A2) in cell A3 and DEPENDS(TIMER.STOP(A3),A1) to time how long the call to the function takes to execute when cell A2 is changed.

The RAND.* are terrific for torture testing spreadsheets. They return random values of all the Excel types that can be used as input to your functions. Start HUEY.RUN and look at memory usage in Task Manager to detect memory leaks.

You can get information about documents, windows, and workspaces with GET_DOCUMENT, GET_WINDOW and GET_WORKSPACE. You will need MACROFUN.HLP to know what arguments to pass and how to interpret the results.

There are also some special purpose functions such as UTILITY.DEPENDS, alias DEPENDS, which is used to control calculation order. The function UTILITY.VOLATILE(Cell), alias VOLATIZE, will make a function call in a Cell volatile. UTILITY.EVALUATE(Expr), alias EVAL, evaluates Expr similar to pressing F9 on selected text in the formula bar. Somewhat more esoteric is UTILITY.CALLER(), alias CALLING. It detects if the user is actively editing a cell. Finally, THIS() returns the previous contents of a cell. One use is to put =A2 + THIS() in cell A1. As you type numbers into cell A2 you will see the running sum in cell A1 (unless you type the same number twice, in which case Excel will "optimize" you spreadsheet and not recalculate A1). If you put IF(CALLING(), 1, A2 + THIS()) in cell A1, then place your cursor in the cell, hit F2, then Enter it will reset the counter to 1. This does not work as expected if you replace 1 by 0. If a cell contains the value 0, then the CALLING() gimmick no longer works. Strange things happen when you try to get Excel to do things it was never designed to do.

Last edited Aug 11, 2013 at 9:05 PM by keithalewis, version 16