## loop structures

The loop structures are an integral part of VBA programming. You are pretty much guaranteed to use them immediately. In some cases, you are more likely to use loops than logic structures. The reason that loops are so critical is that they allow you to perform an action multiple times or across multiple objects. Given the nature of a spreadsheet (where you have a high multitude of cells) and the reasons for using VBA (you want to perform some action multiple times) you really can't avoid loops. Gaining an understanding and comfort with loops is critical to your skill with VBA.

There are several types of loops that work similarly but have different use cases. Those include:

- For Each - useful when you have a collection and want to do something for each object in that collection (this is the most common loop to use since you will nearly always have a collection of Ranges or some other object to iterate through)
- For - useful when you want to do something a specific number of times
- Do/While - run a loop until a condition is met which is useful when you do not know in advance how many times to run the loop and you don't have a finite collection

It is worth noting that all loops can be written as a Do/While loop, but you will nearly never do this. There are good reasons that the For Each and For loops exist.

It is also worth mentioning here that I typically try to avoid For loops whenever possible. I always prefer to use a For Each loop if it is appropriate for the application. This is not an approach that I used from the beginning but have begun to value the use of For Each loops. If you are coming from a programming language that does not value collection iteration, then you might avoid the For Each loop at first. I'd strongly recommend you learn to use the For Each and appreciate it. Your code will be much cleaner and easier to read with For Each loops instead of the alternatives. Especially when dealing with Ranges, it is tempting to iterate through them as a nested For loop. You should really avoid this.

TODO: add an example of a bad For loop

### For Each loop

It is not traditional to start with the For Each instead of the For loop, but I personally use the For Each far more so I'll start there.

The For Each loop is used whenever you have an utterable collection. An utterable collection can come from either the Excel object model or your own code. In general, most of the Excel object model returns an utterable collection. This is especially true for Ranges.

TODO: add a list of utterable collections that can be used here

You are not required to put the variable name in the Next line. I recommend not including the variable unless you have tons of code in the loop and are nesting loops. Typically you will rename the variable and then get a compile time error because the variable names don't match. I've never found the variable name in the Next line to help much.

TODO: add an example of a For Each loop

### For loop

Another style of loop that exists is the "bare" `For` loop. This is one of the simplest loops to understand and control. The idea is simple: iterate through a chunk of code a given number of times. The most common forms of the `For` loop work through a fixed number of iterations. One example is easy: if you want to output the numbers 1 through 10 into a column of cells, you can easily use a For loop to output the number. This is a bad example though since it can easily be done with normal Excel functions, but it is quite common to do the equivalent sort of task when writing a larger macro. In that sense, it is easy to forget how versatile the For loop can be when needing to do something some number of times.

Compared to a While loop (discussed below) there are a number of advantages to the For loop:

- Much easier to control the "exit" strategy and avoid infinite loops
- Can be wired up with constants that are intuitive and just work
- Can be extended to use variables instead of constants to provide more flexibility

When moving beyond the simple 1 to 10 For loop, there are a handful of options which can be pushed into ever complex strategies:

- Use a variable for either the starting index, increment, or end point
- Use a negative step to go backwards through a list of numbers
- Use `Exit For` statements to control execution and kick out of the loop

On that last point is where you will see VBA is woefully underpowered compared to "modern" programming languages. VBA does not provide a simple command to `continue` a loop. There is a `Exit For` which can be used to kick out of the loop, but to `continue` you must create a `:LABEL` and use a `Goto LABEL` statement to jump there. There is nothing necessarily wrong with this, but it is an approach that is annoying and prone to some mistakes. The biggest issue is accidentally moving the label or having some other position issue. THe annoyance is having to create a label and use a `Goto`. There is nothing inherently wrong with a `Goto` but they provide create power which means awful bugs later.

It is worth noting that the `For Each` loop is a simplification of a `For` loop for a number of instances. In most cases, you could create an index and then iterate through an object by index, storing a reference to the object being stored. Behind the scenes, I believe this is how the majority of internal commands are handled. Despite the 1:1 translation between the two loops, it is typically MUCH simpler to use a `For Each` loop if you just need access to the underlying object in the collection. I will nearly always create an index outside of the loop and use it alongside a `For Each` instead of creating a `For` loop and storing a reference to an object. It always seems vastly simpler to store an Integer than an object. Aside from the marginal advantage of not calling `Set`, there is a immediate payoff of a `For Each` loop that if you name the variables correctly, you can typically read exactly what the code will do. This pays dividends for yourself and others later when reviewing the code.

It is also worth mentioning that there are a handful of instances where you are typically required to use a For loop even if you want to use the object being used. The standard example here is if you will eb modified the collection that you are iterating. In this case, you will rapidly create iteration issues trying to modify the collection inside the loop using it. In some of those cases, you will get runtime errors, but in others you will just get unintended consequences. The most common example of doing this is when you want to delete items from a collection while iterating through it. Let's say you need to check whether an item meets some criteria before deleting it. There are two ways to handle this:

- Use a For loop and run thought eh collection BACKWARDS. The direction is critical because it means that at worst, you are working at the end of the list which will not affect future operations.
- Use a "dual loop" approach.

An example of item 1 is shown below.

TODO: add example of backward loop

The dual loop approach is worth mentioning further since sometimes it can give you an elegant way out of a bind. The idea is that instead of modifying the collection while you iterate it, you store some amount of information outside of the collection and then use that to determine what to delete. This typically only works if the items being deleted exist independently of the collection that holds them. This happens often enough with Excel that it is worth giving a concrete example: deleting Rows from a Worksheet. To handle this dual loop approach, there are two possible options:

- Use one loop to create a collection that stores the rows to be deleted and then iterate that collection in a new loop
- Build a larger Range to delete as you go and then use an Excel function to handle the actual deletion

The latter option is only technically a "dual" loop. Technically Excel will use some sort of internal loop to actually delete the Range. You are only required to cleverly create the Range which allows this internal process to be kicked off.

TODO: add an example of the Collection approach for deleting Ranges

TODO: add an example of the UNION-DELETE approach for deleting ranges.

TODO: add some examples of For loops and how they might be used

### Do/While loop

The final style of loop is the Do/While loop. Although it is mentioned last, it is ultimately the simplest type of loop that exists. The idea is: run until some condition is meeting. This loop matches very nicely when your looping strategy involves some condition. You simply put than tocndifiton int he loop and let it work. The downside to a Do?while loop comes down to the possibility of an infinite loop. This leads to the common problem of a macro that hangs Excel and requires intervention to shut down. Infinite loops are technically easy to avoid, but it is far more common in practice to skip the steps that help avoid infinite loops.

It is worth mentioning at this point that all of the loop varieties can be recreated from the other loop varieties. From this standpoint, there are slight advantages to one style over another, but at the end of the day, you simply write the loop that works for the task at hand.

For a Do/While loop, there are two possible ways of writing it. You can either do a `Do...While` or a `While...Loop`. THe main difference is whether or not the loop will execute before the condition is checked. There are instances where one style makes more sense over the other. Typically you can always use the `While...Loop` variety, but you may be required to type an initialization statement before the loop that is repeated within the loop.

Some common examples where a While loop make sense include:

- Iterate down through a column of cells until some condition is met (typically a blank or non-blank cell). This is quite helpful when it is difficult to create the `Range` that might be used for a `For Each` loop.
- Iterate through the file system using the `Dir` command to find files to open and process

The WHile loop tends to make the most sense when you are not iterating through a fixed collection of objects because the `For Each` does a better job there. You also would avoid using it when you have a fixed number of iterations to run where a `For` loop makes a lot more sense. That then leaves the instances where you want to loop through some action some number of times, but you're not sure how many times until you start going.

If you are particular adventurous, you can make use of the `Exit Do` command to exit out of the loop mid iteration. This pairs nicle with a `While True` at the start of the loop to ensure that nothing else will kick you out of the loop. There are instances where this can be a simple way to loop, but you have to be absolutely certain your `Exit Do` command will be triggered at some ppint or else you guarantee an infinite loop.

TODO: add an example of looping use `Dir`

TODO: add an example of a loop that works through a range using Offset

### which loop and why

There are a handful of common reasons you might go for one loop instead of another. Worth knowing is that in general you can use any of the loops and get the same result. One approach is typically much easier to program, understand, and maintain.

A couple of good things to remember:

If you are going to modify a collection in the course of iterating through it, you should not use a For Each loop. The For Each does not update the utterable collection if you modify it during a loop. This is particularly important if you are looping through a collection to identify items to delete from the collection. You should never do this in a For Each loop. When deleting, you should typically use a For loop and iterate through the collection in reverse order. This makes it easy to handle deleting items since you cannot get out of order. You can achieve the same result with a Do/While, but I won't cover that.

TODO: add an example of deleting via a For loop

If you need an index/incrementing variable alongside your loop, you are not required to use a For loop. You can always create a new variable and increment it yourself inside the loop. This is sometimes preferable to switching to a For loop solely to get the counter/index variable.

If you are using a Do/While loop, you should give serious consideration to adding a counter and breaking the loop if the counter gets too large. It happens far too often where I use a While loop and end up freezing Excel because the loop never terminates. You can sometimes break the code and get Excel to respond, but that does not always work. This is especially important if you are generating code that others will use since they may be less familiar with how to break out of an infinite loop.

You may need to break out of a loop. Unfortunately, VBA does not have the normal Break and Continue commands that you might be familiar with from another language. The only way to break out of a loop on the spot is to add a label use a Goto command unless you are able to break out of the Function/Sub completely using Exit. This always feels dirty to me so instead I will typically structure the loop with a Boolean that can detect whether the next iteration should continue. This works for Continue, but it is not a good solution for getting a Break. The only way to do this is via a Goto. Just do it.
