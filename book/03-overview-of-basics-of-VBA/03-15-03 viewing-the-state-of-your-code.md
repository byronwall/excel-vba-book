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
