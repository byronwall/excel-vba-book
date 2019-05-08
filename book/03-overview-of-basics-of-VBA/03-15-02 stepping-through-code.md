### stepping through code

Once you have entered the debugger, tehre are a handful of ways to affect execution. They are:

- Run
- Step Into
- Step Over

TODO: add a picture of the toolbar icons

TODO: explain how to reach these comamdns along with the shortcuts

Run will tell teh debugger to just keep running until it hits another error or breakpoint. This is the same as normal execution.

Step Into and Step Over do the same thing with one difference. They both tell VBA to execute the current instruction adn then resume debugging after it. The difference is how they hadnle whether or not to enter a `Sub` or `Function`. If you have a written a Sub or Functojn of your own adn then call it, you ahve tow options while debugging. You can either enter that Sub and step through the commands in tehre. Or, you can treat that line with teh Sub as a single step which can be processed as a single instruction. If you do that, you will `Step Over` all of the intermediate execution and reusme deugging once code returns back to the level you started at. This is very important if you ahve a large number of nested Subs and Functiojns. The debugging steps allow you to decided how "deep" into the call stack you will go to pursue your deugggin. Soemtimes, you will know that a givne Sub works as intended and you do not want to step into it. Other times, you will reach a Sub being called and want to know exactly how it arrived at its output.

If you want to step through to a specific spot but cannot get there easily with the commands above, you can always just set a new breakpoint right there and hit `Run`. This will run until that line. You can also right click on a line nad do `Run until this point` and you will get the same effect. TODO: is that right?
