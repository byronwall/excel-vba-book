### viewing the state of your code

The whole point of debugging is generally to view the state of oyur code (or the Excel side of things) in process.  The idea of viewing teh state menaes a couple of concrete things:

* What are the values of specific variables?
* What was the order of execution?  Which control structures were processed and in what way?
* What happens if I do "this" instead of "that"?

Each of those is hit below:

#### values of variables

Typically, the most important aspect of debugging is seeing whcih variables hold which values.  The idea is taht if you can see waht the variables hold at runtime, you can cehck that against your expecatations and then gain insight into why your program is behaving the way it does.  Other times, you want to see the values of things so that you can decide how to proceed from your current ppint.  VBA provides a number of ways to check the value of a variable:

* Hover over teh variable and allow the VBE to see you the value
* Using the Locals window
* Using the Immediate window with `?` added to the start (TODO: is that the same as Debug.pRint?)
* Using the Watch window after creating a watch
* Running a command where you put the value into the spreadsheet

The VBE is fairly helpful when debgging comapred to other debuggers.  It does about what you would expect.  This means that you will get tooltips when you hover over variables.  This works well for variables that hold a value and not an object. For an object, if you hover, you will get the `.Value` property of the object and not a drop down to explore. IN this regard, the debugger is inferior to a modern Visual Studio instance.

If you want to explore the properties of an object, or see a persistent value without hovering, you can use the Loacls or Watch widnow.  They do the same thing: show the values of variables while also alllowing you to click down into Objects and their properties.  The Locals windwo works by giving you a list of all the local variables automatically. T eh Watch window works by requirng you to provide the vairable name or caommnad that you watn to watch. I always start with teh Locals window since tpyically local variabel are what I want to see.

When reviewing the contents of an object, beware that VBA will not show you all of the properties of the object.  In particular, it will not show you properties that are the result of a Function instead of a nomral property.  For a lot of Excel Object Model objects this is a key point.  THere are a large nujmber of properties that you will need to add ot the Watch window or query directly with teh Immediate window to see their value.  A common example: `Range.Address`.

TODO: add an example of using the Watch window

TO use the Immediate iwndow, you first need to enable it via View (TODO: add this for others).  Once enabled, you can use the Immediate window as a palce to execute whatever code you want.  It works by executing single lines at a time.  IF you want the output of a command, use `?` at the start to print the result.  You can use the Immeidate window whevnerm, including duirng normal development (i.e. even when code is not running).

TODO: add an example of using the Immedaite window

One particular thing that can be done (although not often) is that you can use the spreadsheet as a place to dump the results of oyur debugging.  Sometimes, you will need to inspect some oibject and find that the VBE is just not tha thelpful.  Maybe you have an array whose values you want to hceck.  The simpple approach here is to dump that array to the spreadsheet using teh Immediate windwo (or actual code) and then set a breakppint to inspect it.  This igves a nice back and forth between Excel and VBA that simply does not exist in other programming environments.  ONce you see Excel as a huge playgournd to dump arrays, you will find all sorts of using for that while programming.
