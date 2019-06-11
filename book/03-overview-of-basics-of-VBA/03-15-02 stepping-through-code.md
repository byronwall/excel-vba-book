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
