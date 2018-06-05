### For loop

Another style of loop that exists is the "bare" `For` loop.  This is one of the simplest loops to understand and control.  The idea is simple: iterate through a chunk of code a given number of times.  The most common forms of the `For` loop work through a fixed number of iterations.  One example is easy: if you want to output the numbers 1 through 10 into a column of cells, you can easily use a For loop to output the number.  This is a bad example though since it can easily be done with normal Excel functions, but it is quite common to do the equivlanet sort of task when writing a larger macro.  In that sense, it is easy to forget how versatile the For loop can be when needing to do something some number of times.

Compared to a While loop  (discussed below) there are a number of advantages to the For loop:

- Much easier to control the "exit" strategy and avoid infinite loops
- Can be wired up with constants that are intutiive and just work
- Can be extended to use variables instead of constants to provide more flexibiilty

When moving beyond the simple 1 to 10 For loop, there are a handful of options which can be pusehd into ever complex strategies:

- Use a variable for either the starting index, increment, or end point
- Use a negative step to go backwards through a list of numbers
- Use `Exit For` statements to control execution and kick out of the loop

On that last point is where you will see VBA is woefully underppowered compared to "modern" programming lanugages.  VBA does not provide a simple command to `continue` a loop.  There is a `Exit For` which can be used to kick out of the loop, but to `continue` you must create a `:LABEL` and use a `Goto LABEL` statement to jump there.  There is nothign necessarily worng iwth this, but it is an approach that is annoying and prone to some mistakes.  The biggest issue is accidentally moving the label or having soem other positon issue.  THe annoyance is having to create a label and use a `Goto`.  There is nothing inherently wrong with a `Goto` but they provide create power which means awful bugs later.

It is worht noting that the `For Each` loop is a simplification of a `For` loop for a number of instances.  In most cases, you coudl create an index and then iterate through an object by index, storing a reference to the object being stored.  Behidn the scenes, I believe this is how the majority of internal commands are hanlded.  Depsite the 1:1 translation between the two loops, it is typically MUCH simpler to use a `For Each` loop if you just need access to the underlying object in the collection.  I will nearly always create an index outside of the loop and use it alongside a `For Each` instead of creating a `For` loop and storing a reference to an object.  It always seems vastly simpler to store an Integer than an object.  Aside from the marginal advantage of not calling `Set`, there is a immediate payoff of a `For Each` loop that if you name the variables correctly, you can typically read exaclty what the code will do.  This pays dividends for yourself and others later when reviewining the code.

It is also worth metnioning that there are a handful of instances where you are typically required to use a For loop even if you want to use the object being used.  The standard example here is if you will eb modified the colleciton that you are iterating.  In this case, you will rapidly create iteration issues trying to modify the collecion inside the loop using it.  In some of those cases, you will get runtime errors, but in others you will just get unintended consequences.  The most common example of doing this is when you want to delete items from a colleciton while iterating through it.  Let's say you need to check whether an item meets some critertia before deleting it.  There are two ways to handle this:

- Use a For loop adn run throught eh collection BACKWARDS.  The direction is crtiical because it means that at wrost, you are working at the end of the list which will nto affect future operations.
- Use a "dual loop" approach.

An example of item 1 is sjown below.

TODO: add example of backward loop

The dual loop approach is worth mentionign further since sometimes it can give you an elegant way out of a bind.  The idea is that instead of modifying the colleciton while you iterate it, you store some amount of information outside of the collection and then use that to determine what to delete.  This typically only works if the items being deleted exist independetly of the colleciton that holds them.  This happens often enough with Excel that it is worht giving a concrete example: deleting Rows from a Worksheet.  To handle this dual loop approach, there are two possible options:

- Use one loop to create a colleciton that stores the rows to be deleted and then iterate that collection in a new loop
- Build a larger Range to delete as you go and then use an Excel funciton to handle the actual deletion

The latter option is only technically a "dual" loop.  Tehcnically Excel will use some sort of internal loop to actually delete the Range.  You are only required to cleverly create the Range which allows this internal process to be kicked off.

TODO: add an example of the Collection approahc for deleting Ranges

TODO: add an example of the UNION-DELETE approach for dleeting ranges.

TODO: add some examples  of For loops and how they might be used
