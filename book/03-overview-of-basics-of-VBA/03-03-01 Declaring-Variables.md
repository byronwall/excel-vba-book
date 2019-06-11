### Declaring Variables

Declaring variables is a straight forward tasks. VBA offers a simple command to declare a new variable: `Dim`. Note that you can technically use a variable before declaring it, but you should really avoid this practice. It leads to the potential to create all sorts of bugs later. Just don't do it. To better avoid this, setting the flag in the settings (TODO: add a picture of that).

When declaring a variable, there are two components to it: variable name and variable type. Variable name is wholly your decision with only a couple of constraints. You are not allowed to duplicate the name of an internal command, and you should go to some length to avoid using the same name as an Excel object model name. Beware that naming a variable ahs certain conventions, but these do not have any effect on the program execution. The main concern with names is that they will directly affect your ability to work with and maintain your code. Naming things is hard. Pick a strategy that works for you and your coworkers and get on it with it. There is no single answer here about how to name things.

The second part of the puzzle is to declare the type of the variable. This is THE core part of variables. When declaring a variable, you are essentially deciding if the type should be the generic `Variant` or if you should actually declare a type. Note that there are times when you have to use Variant, but in general, you should use the most specific type that is possible. These types can either draw from VBA or from the Object Model, or from your own created types. When thinking of variable types, there are two major groups of types:

- Value types = a number, string, or boolean
- Reference types = objects
