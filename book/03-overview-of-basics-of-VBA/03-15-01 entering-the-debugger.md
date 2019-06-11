### entering the debugger

To enter the debugger, you need to either set a breakpoint, hit Step Into, hit the Break key, or have an error thrown that prompts for debugging. By default, you will not be using the debugger while your code is running. This si actually a good thing since debuggin code adds a large overhead which will kill performance. The most common approaches to entering the debugger are to set a breakpoint or via an error. This lines up with the idea that you either want to debug a specific point in your code or that you want to be able to see what wnet wrong when an error is thrown.

When setting a breakpoint, tehre are a handful of reasons for choosing where to set one:

- Right before an important step so that you can see the before and after state
- Inside of a control structure so that you can see whether or execution enters that structure. Sometimes there is information to be had when the code does _not_ reach a breakpoint.

When breakpoints, you can technically disable them instead of removing them if you do not want them to tirgger. I never use that feature.

If you are entering the debugger through an error, you simply hit `Debug` on the prompt. You will be starting on the line that threw the error ready to execute it again.

The other ways to enter the debugger are by hitting the CTRL+BREAK shortcut. If the VBA is at a stoppable point, this will cause an interrupt which gives the same prompt as the error prompt. From here, you can hit `Debug`.

The final approach is to use the Step Into button on the code to run. TODO: is this true?
