## debugging UDFs

Debugging a UDF is really teh same as debuggin normal code except you need to understand when your code will be called and hence, what you may be deugging.  The simplest way to think about deugging a UDF is with an empty spreadsheet.  In this example, once you type your UDF into the spreadhseet, Excel will execute the code and you can debug it via a breakppint.  This is simple.

For a larger spreadsheet however, you are very likely to use your UDF more than once while only haivng a problem with a specific instance of it.  Let's say your UDF does some fancy statistics btu cannot handle certin types of intputs.  You can see that your code is throwing an error with a `#VALUE!` output.  If you add a breakpoint to the UDF, then you risk having to debug a large number of successful calls before your bad one happens.

There are a couple of approaches to deal with this:

* Edit a formula for the cell you want with a breakpoint set in the debugger.  Excel will execute that "new" formula first which will be the one of interest.
* Right a quick If statement to check if the teh Caller's address is a specific cell.

The first example is easy enough to understand and si the typical approach for debuggin a UDF.  It's a bit of a pain because your breakpoint will stay in place and may be hit several times later.  To get around this, you can siwtch over to manual calculation to avoid all the other cells calcualting.  TODO: is that right?

The second approach works well when you have a UDF in several place but where only one of them is causing an erorr.  You can add a temporary statement at the top to check for the Calller address and then set a breakpoint inside there.  ocne it's hit, you know you are debugging the right call and can then step through teh code.  You can do the same approach to check for the incoming value or really anythign else that is unique to the problematic cell.  The nice thing ehre is that if you can figure out what statement to use for the breakpoint, you will ahve n aidea of which condiditons may cause the problem.

TODO: how are runtime errors handled here?  any way to get them thrown with a prompt.
