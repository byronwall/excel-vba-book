## debugging UDFs

Debugging a UDF is really the same as debugging normal code except you need to understand when your code will be called and hence, what you may be debugging. The simplest way to think about debugging a UDF is with an empty spreadsheet. In this example, once you type your UDF into the spreadsheet, Excel will execute the code and you can debug it via a breakpoint. This is simple.

For a larger spreadsheet however, you are very likely to use your UDF more than once while only having a problem with a specific instance of it. Let's say your UDF does some fancy statistics btu cannot handle certain types of inputs. You can see that your code is throwing an error with a `#VALUE!` output. If you add a breakpoint to the UDF, then you risk having to debug a large number of successful calls before your bad one happens.

There are a couple of approaches to deal with this:

- Edit a formula for the cell you want with a breakpoint set in the debugger. Excel will execute that "new" formula first which will be the one of interest.
- Right a quick If statement to check if the the Caller's address is a specific cell.

The first example is easy enough to understand and si the typical approach for debugging a UDF. It's a bit of a pain because your breakpoint will stay in place and may be hit several times later. To get around this, you can switch over to manual calculation to avoid all the other cells calculating. TODO: is that right?

The second approach works well when you have a UDF in several place but where only one of them is causing an error. You can add a temporary statement at the top to check for the Caller address and then set a breakpoint inside there. once it's hit, you know you are debugging the right call and can then step through the code. You can do the same approach to check for the incoming value or really anything else that is unique to the problematic cell. The nice thing ehre is that if you can figure out what statement to use for the breakpoint, you will have n aidea of which conditions may cause the problem.

TODO: how are runtime errors handled here? any way to get them thrown with a prompt.
