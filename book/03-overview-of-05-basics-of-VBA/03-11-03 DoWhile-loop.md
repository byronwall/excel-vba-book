### Do/While loop

The final style of loop is the Do/While loop.  Although it is mentioned last, it is ultimately the simplest type of loop that exists.  The idea is: run until some condition is meeting.  This loop matches very nicely when your looping strategy involves some condition.  You simply put tha tocndifiton int he loop adn let it work. The downside to a Do?while loop comes down to the possibility of an infinte loop. This leads to the common problem of a macro that hangs Excel and requires intevention to shut down.  Infinite loops are technically easy to avoid, but it is far more common in practice to skip the steps that help avoid infintie loops.

It is worth mentioning at this point that all of the loop varieties can be recreated from teh other loop varieties.  From this standpoint, there are slight advantages to one style over anotehr, but at the end of hte day, you simply write the loop that works for the task at hand.

For a Do/While loop, there are two possible ways of writing it.  You can either do a `Do...While` or a `While...Loop`.  THe main difference is whether or not the loop will execute before the condition is checked.  There are instances where one style makes more sense over the other.  Typically you can always use the `While...Loop` variety, but you may be required to type an intialization statement before the loop that is repeated within the loop.

Some common examples where a While loop make sense include:

- Iterate down through a column of cells until some condition is met (typically a blank or non-blank cell).  This is quite helpful when it is difficult to create the `Range` that might be used for a `For Each` loop.
- Iterate thruogh the file system usign the `Dir` command to find files to open and process

The WHile loop tends ot make the most sense when you are not iterating through a fixed colleciton of objects bcause the `For Each` does a better job there.  You also would avoid using it when you have a fixed nujmber of iterations to run where a `For` loop makes a lot more sense.  That then leaves the instances where you want to loop through some action some nujmber of times, but you're not sure how many times untli you start going.

If you are particualr adventurous, you can make use of the `Exit Do` command to exit out of the loop mid iteration.  This pairs nicle with a `While True` at the start of hte loop to ensure that nothign else will kick you out of hte loop.  There are instances where this can be a simple way to loop, but you have to be absolutely certain your `Exit Do` command will be triggedered at some ppint or else you guarantee an infinite loop.

TODO: add an example of looping use `Dir`

TODO: add an example of a loop that works through a range using Offset
