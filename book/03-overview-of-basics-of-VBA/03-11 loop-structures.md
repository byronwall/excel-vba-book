## loop structures

The loop structures are an integral part of VBA programming. You are pretty much guaranteed to use them immediately. In some cases, you are more likely to use loops than logic structures. The reason that loops are so critical is that they allow you to perform an action multiple times or across multiple objects. Given the nature of a spreadsheet (where you have a high multitude of cells) and the reasons for using VBA (you want to perform some action multiple times) you really can't avoid loops. Gaining an understanding and comfort with loops is critical to your skill with VBA.

There are several types of loops that work similarly but have different use cases. Those include:

- For Each - useful when you have a collection and want to do something for each object in that collection (this is the most common loop to use since you will nearly always have a collection of Ranges or some other object to iterate through)
- For - useful when you want to do something a specific number of times
- Do/While - run a loop until a condition is met which is useful when you do not know in advance how many times to run the loop and you don't have a finite collection

It is worth noting that all loops can be written as a Do/While loop, but you will nearly never do this. There are good reasons that the For Each and For loops exist.

It is also worth mentioning here that I typically try to avoid For loops whenever possible. I always prefer to use a For Each loop if it is appropriate for the application. This is not an approach that I used from the beginning but have begun to value the use of For Each loops. If you are coming from a programming language that does not value collection iteration, then you might avoid the For Each loop at first. I'd strongly recommend you learn to use the For Each and appreciate it. Your code will be much cleaner and easier to read with For Each loops instead of the alternatives. Especially when dealing with Ranges, it is tempting to iterate through them as a nested For loop. You should really avoid this.

TODO: add an example of a bad For loop
