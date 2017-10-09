## logic structures

The logic structures are the backbone of nearly all VBA programs.  There are a handful of times where you can just run through some commands with no branching logic, but in general, your program will need to make decisions based on some condition that it encounters.  In order to make those decisions, you use the logic structures of VBA.  There is really only a single logic structure in VBA, the If-Then structure, but VBA provides a handful of useful additions to the basic If-Then to make programming a little easier.

The main logic structures then are:

* If-Then
* If-ElseIf-Then
* Select-Case

The If-Then is the main building block that allows you to do something if a condition is true or do something else otherwise.  The If-ElseIf-Then allows you to add additional conditions to check before defaulting to the Else.  If any of the ElseIf statements evaluate True, then the branch will stop traversing the conditions.  The Select-Case is an extension of the If-ElseIf-Then that always compares a given variable against different possible values.

TODO: add example of the different forms of If-Then

For logic evaluation, there are always a handful of ways to arrive at the same result.  You are allowed to evaluate multiple conditions in a single If-Then statement by using And and Or.  You can also "nest" different logic blocks inside of each other to create the same sort of logic.  In this way, you can either use If-ElseIf-Then or you can do an If-Then and stick a second If-Then in the Else clause.  These will be equivalent.  Sometimes one version looks better or makes more sense than the other.

## helpful logic functions

TODO: add an overview of the functions And/Or (is there a XOR?) along with the logic operators <,>,<>, = etc.

## the Select Case

The Select Case makes it possible to compare the value of a variable against multiple values without having to type the variable name every time.  This is a purely syntactic feature that makes programming easier in certain cases.  You can completely duplicate a Select-Case with an If-ElseIf-Then, but you may have to type more code.

TODO: add an example of a Select Case

Note that this section is fairly short.  Logic structures are so prevalent in normal VBA code that should look for examples of these in the respective chapters instead of this section.
