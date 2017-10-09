## slightly more complicated manipulations (the two steppers)

This section will on the so called "two steppers".  I call them that because these manipulations typically involve two commands after identifying a `Range`.  the first command is usually a logic or loop, and the second command is the actual work ot be done.  Two steppers are important because a large number of complicated tasks involve nesting and combining these two steps.

Some examples of two step  manipulations includes;

* Run through a list of cells, if the text is numeric, convert to a number
* Run through a list of cells, if the cell is blank, fill with the value from above
* Run through some cells, check if the row is odd or even, and color the row from one of two colors
* Run through one list of cells, apply the formatting to the same cell in a different column

TODO: find some better examples for these as well

### strategy #1, do something if

This strategy really is the core of all advanced VBA development.  It's simple enough: "do something, if".  The endless possibilities come from the choices for "do something" and the things that could be checked in the "if".  There are a handful of common scenarios that are best covered by storing some utility code (e.g. convert to a number if numeric).  Most of these two step solutions though are specific to the task at hand.

In this section, the goal is to show the general form of this strategy with a couple of examples.

TODO: add a couple examples of this

### strategy #2, work through one `Range` and apply to another `Range`

This strategy comes up frequently when working through `Ranges` that are related somehow.  The general idea is that you want to apply an action in one `Range` based on something about another `Range`.  The simplest case of this is to move a value from `Range` to another.  This simple case sometimes reduces to not much more than copying and pasting.  Having said that, once you get past the simplest version of it, you will be doing something that copy and paste cannot handle.

TODO: add a couple examples of this
