### which loop and why

There are a handful of common reasons you might go for one loop instead of another. Worth knowing is that in general you can use any of the loops and get the same result. One approach is typically much easier to program, understand, and maintain.

A couple of good things to remember:

If you are going to modify a collection in the course of iterating through it, you should not use a For Each loop. The For Each does not update the iterable collection if you modify it during a loop. This is particularly important if you are looping through a collection to identify items to delete from the collection. You should never do this in a For Each loop. When deleting, you should typically use a For loop and iterate through the collection in reverse order. This makes it easy to handle deleting items since you cannot get out of order. You can achieve the same result with a Do/While, but I won't cover that.

TODO: add an example of deleting via a For loop

If you need an index/incrementing variable alongside your loop, you are not required to use a For loop. You can always create a new variable and increment it yourself inside the loop. This is sometimes preferable to switching to a For loop solely to get the counter/index variable.

If you are using a Do/While loop, you should give serious consideration to adding a counter and breaking the loop if the counter gets too large. It happens far too often where I use a While loop and end up freezing Excel because the loop never terminates. You can sometimes break the code and get Excel to respond, but that does not always work. This is especially important if you are generating code that others will use since they may be less familiar with how to break out of an infinite loop.

You may need to break out of a loop. Unfortunately, VBA does not have the normal Break and Continue commands that you might be familiar with from another language. The only way to break out of a loop on the spot is to add a label use a Goto command unless you are able to break out of the Function/Sub completely using Exit. This always feels dirty to me so instead I will typically structure the loop with a Boolean that can detect whether the next iteration should continue. This works for Continue, but it is not a good solution for getting a Break. The only way to do this is via a Goto. Just do it.
