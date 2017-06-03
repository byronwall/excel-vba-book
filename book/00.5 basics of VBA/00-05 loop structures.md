## loop structures

The loop structures are an integral part of VBA programming.  You are pretty much guaranteed to use them immediately.  In some cases, you are more likely to use loops than logic structures.  The reason that loops are so critical is that they allow you to perform a action multiple times or across multiple objects.  Given the nature of a spreadsheet (where you have a high multitude of cells) and the reasons for using VBA (you want to perform some action multiple times) you really can't avoid loops.  Gaining an understanding and comfort will loops is critical to your skill with VBA.

There are several types of loops that work similarly but have different use cases.  Those include:

* For Each loop - useful when you have a collection and want to do something for each object in that collection (this is the most common loop to use since you will nearly always have a collection of Ranges or some other object to iterate through)
* For loop - useful when you want to do something a specific number of times
* Do/While loops - run a loop until a condition is met which is useful when you do not know in advance how many times to run the loop and you don't have a finite collection

It is worth noting that all loops can be written as a Do/While loop, but you will nearly never do this.  There are good reasons that the For Each and For loops exist.

It is also worth mentioning here that I typically try to avoid For loops whenever possible.  I always prefer to use a For Each loop if it is appropriate for the application.  This is not an approach that I used from the beginning but have begun to value the use of For Each loops.  If you are coming from a programming language that does not value collection iteration, then you might avoid the For Each loop at first.  I'd strongly recommend you learn to use the For Each and appreciate it.  Your code will be much cleaner and easier to read with For Each loops instead of the alternatives.  Especially when dealing with Ranges, it is tempting to iterate through them as a nested For loop.  You should really avoid this.

TODO: add an example of a bad For loop

### For Each loop

It is not traditional to start with the For Each instead of the For loop, but I personally use the For Each far more so I'll start there.

The For Each loop is used whenever you have an iterable collection.  An iterable collection can come from either the Excel object model or your own code.  In general, most of the Excel object model returns an iterable collection.  This is especially true for Ranges.

You are not required to put the variable name in the Next line.  I recommend not including the variable unless you have tons of code in the loop and are nesting loops. Typically you will rename the variable and then get a compile time error because the variable names don't match.  I've never found the variable name in the Next line to help much.

TODO: add an example of a For Each loop

### For loop

TODO: add content

### Do/While loop

TODO: add content

### which loop and why

There are a handful of common reasons you might go for one loop instead of another.  Worth knowing is that in general you can use any of the loops and get the same result.  One approach is typically much easier to program, understand, and maintain.

A couple of good things to remember:

If you are going to modify a collection in the course of iterating through it, you should not use a For Each loop.  The For Each does not update the iterable collection if you modify it during a loop.  This is particularly important if you are looping through a collection to identify items to delete from the collection.  You should never do this in a For Each loop.  When deleting, you should typically use a For loop and iterate through the collection in reverse order.  This makes it easy to handle deleting items since you cannot get out of order.  You can achieve the same result with a Do/While, but I won't cover that.

TODO: add an example of deleting via a For loop

If you need an index/incrementing variable alongside your loop, you are not required to use a For loop.  You can always create a new variable and increment it yourself inside the loop.  This is sometimes preferable to switching to a For loop solely to get the counter/index variable.

If you are using a Do/While loop, you should give serious consideration to adding a counter and breaking the loop if the counter gets too large.  It happens far too often where I use a While loop and end up freezing Excel because the loop never terminates.  You can sometimes break the code and get Excel to respond, but that does not always work.  This is especially important if you are generating code that others will use since they may be less familiar with how to break out of an infinite loop.

You may need to break out of a loop.  Unfortunately, VBA does not have the normal Break and Continue commands that you might be familiar with from another language.  The only way to break out of a loop on the spot is to add a label use a Goto command.  This always feels dirty to me so instead I will typically structure the loop with a Boolean that can detect whether the next iteration should continue.  This works for Continue, but it is not a good solution for getting a Break.  The only way to do this is via a Goto.  Just do it.
