## debugging your VBA code

One of the most useful features of VBA and the VBE is the ability to debug your code simply and in place. It is easy to take for granted the power of the VBE debugger, but it is worth mentioning that it is a solid debugger. The debugger has a handful of specific uses related to debugging your code:

- Stepping through execution and watching the movement of values into and out of variables
- Using the Immediate Window to execute arbitrary code or output the results of some value
- Setting the next instruction to force VBA to jump to an arbitrary point in your code
- Viewing the call stack to see how you reached a given spot
- Breaking at an arbitrary breakpoint or after an error was thrown

### entering the debugger

To enter the debugger, you need to either set a breakpoint, hit Step Into, hit the Break key, or have an error thrown that prompts for debugging. By default, you will not be using the debugger while your code is running. This si actually a good thing since debugging code adds a large overhead which will kill performance. The most common approaches to entering the debugger are to set a breakpoint or via an error. This lines up with the idea that you either want to debug a specific point in your code or that you want to be able to see what went wrong when an error is thrown.

When setting a breakpoint, there are a handful of reasons for choosing where to set one:

- Right before an important step so that you can see the before and after state
- Inside of a control structure so that you can see whether or execution enters that structure. Sometimes there is information to be had when the code does _not_ reach a breakpoint.

When breakpoints, you can technically disable them instead of removing them if you do not want them to trigger. I never use that feature.

If you are entering the debugger through an error, you simply hit `Debug` on the prompt. You will be starting on the line that threw the error ready to execute it again.

The other ways to enter the debugger are by hitting the CTRL+BREAK shortcut. If the VBA is at a stoppable point, this will cause an interrupt which gives the same prompt as the error prompt. From here, you can hit `Debug`.

The final approach is to use the Step Into button on the code to run. TODO: is this true?

### stepping through code

Once you have entered the debugger, there are a handful of ways to affect execution. They are:

- Run
- Step Into
- Step Over

TODO: add a picture of the toolbar icons

TODO: explain how to reach these commands along with the shortcuts

Run will tell the debugger to just keep running until it hits another error or breakpoint. This is the same as normal execution.

Step Into and Step Over do the same thing with one difference. They both tell VBA to execute the current instruction and then resume debugging after it. The difference is how they handle whether or not to enter a `Sub` or `Function`. If you have a written a Sub or Function of your own and then call it, you have tow options while debugging. You can either enter that Sub and step through the commands in there. Or, you can treat that line with the Sub as a single step which can be processed as a single instruction. If you do that, you will `Step Over` all of the intermediate execution and resume debugging once code returns back to the level you started at. This is very important if you have a large number of nested Subs and Functions. The debugging steps allow you to decided how "deep" into the call stack you will go to pursue your deign. Sometimes, you will know that a given Sub works as intended and you do not want to step into it. Other times, you will reach a Sub being called and want to know exactly how it arrived at its output.

If you want to step through to a specific spot but cannot get there easily with the commands above, you can always just set a new breakpoint right there and hit `Run`. This will run until that line. You can also right click on a line nad do `Run until this point` and you will get the same effect. TODO: is that right?

### viewing the state of your code

The whole point of debugging is generally to view the state of your code (or the Excel side of things) in process. The idea of viewing the state means a couple of concrete things:

- What are the values of specific variables?
- What was the order of execution? Which control structures were processed and in what way?
- What happens if I do "this" instead of "that"?

Each of those is hit below:

#### values of variables

Typically, the most important aspect of debugging is seeing which variables hold which values. The idea is that if you can see what the variables hold at runtime, you can check that against your expectations and then gain insight into why your program is behaving the way it does. Other times, you want to see the values of things so that you can decide how to proceed from your current ppint. VBA provides a number of ways to check the value of a variable:

- Hover over the variable and allow the VBE to see you the value
- Using the Locals window
- Using the Immediate window with `?` added to the start (TODO: is that the same as Debug.pRint?)
- Using the Watch window after creating a watch
- Running a command where you put the value into the spreadsheet

The VBE is fairly helpful when debugging compared to other debuggers. It does about what you would expect. This means that you will get tooltips when you hover over variables. This works well for variables that hold a value and not an object. For an object, if you hover, you will get the `.Value` property of the object and not a drop down to explore. IN this regard, the debugger is inferior to a modern Visual Studio instance.

If you want to explore the properties of an object, or see a persistent value without hovering, you can use the Loacls or Watch window. They do the same thing: show the values of variables while also allowing you to click down into Objects and their properties. The Locals window works by giving you a list of all the local variables automatically. T eh Watch window works by requiring you to provide the variable name or command that you want to watch. I always start with the Locals window since typically local variable are what I want to see.

When reviewing the contents of an object, beware that VBA will not show you all of the properties of the object. In particular, it will not show you properties that are the result of a Function instead of a normal property. For a lot of Excel Object Model objects this is a key point. There are a large number of properties that you will need to add to the Watch window or query directly with the Immediate window to see their value. A common example: `Range.Address`.

TODO: add an example of using the Watch window

TO use the Immediate window, you first need to enable it via View (TODO: add this for others). Once enabled, you can use the Immediate window as a place to execute whatever code you want. It works by executing single lines at a time. IF you want the output of a command, use `?` at the start to print the result. You can use the Immediate window whevnerm, including during normal development (i.e. even when code is not running).

TODO: add an example of using the Immediate window

One particular thing that can be done (although not often) is that you can use the spreadsheet as a place to dump the results of your debugging. Sometimes, you will need to inspect some object and find that the VBE is just not than helpful. Maybe you have an array whose values you want to hceck. The simple approach here is to dump that array to the spreadsheet using the Immediate window (or actual code) and then set a breakpoint to inspect it. This gives a nice back and forth between Excel and VBA that simply does not exist in other programming environments. Once you see Excel as a huge playgournd to dump arrays, you will find all sorts of using for that while programming.

### forcing execution

In addition to watching the execution of a program, you also have the ability to change the execution. T his si done by using the `Set next command` TODO: name? while running. This is the "nuclear" option of debugging because it does exactly what it says. It will tell VBA to execute _whatever_ line you want next. This allows you to completely ruin your execution while also providing you the power to step to a given spot. It's always the case when writing code that you end up on the wrong side of an If/Else while developing a loop. Sometimes, you just want to see what happens if you go down the other branch. T his option allows you to test that alternative execution path without having to modify your code. You just tell the debugger to execute that branch next and things will work. Very often however, you can using this feature to accidecanlly skip over code where variables are declared or Set and then you will have all sorts of errors because objects are set to `Nothing` instead of the values that are required.

Despite the pitfalls of moving execution arbitrarily, most people who know this feature exists are capable of using it appropriately. They typically are not surprised when things break.

### viewing the call stack

One final feature which is useful is to check the Call Stack. The Call Stack is a list of all the proceduring Subs or Functions that are "active" preceding the current command. It gives you a list of all the places that came before your current line of code. The Call Stack is invaluable when you have started debugging following an error because oftentimes you will not know how you reached a given spot. This is especially true if you are debugging code that is used in multiple places.

To see the Call Stack do View->Call Stack. You can then double click on an item and jump back to that spot. Note that the VBE will attempt to show you the vales of variables at that location which can be very helpful.

The Call Stack can be very helpful if you are using recursive code that calls itself. This code can be very hard to debug because oftentimes a breakpoint will trigger more than you want. IF you are waiting for an error on the 8th time through a Function, then you don't want to skip the breakpoint 7 times. Instead, you can wait for the error, then use the Call Stack to step back through the previous iterations and see what happened.

TODO: add a picture of the call stack
