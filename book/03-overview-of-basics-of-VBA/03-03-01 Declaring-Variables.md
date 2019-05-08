### Declaring Variables

Decalring vairables is a straight forward tasks. VBA offers a simple command to declare a new variable: `Dim`. Note that you can technically use a vriable before declaring it, but you should really avaoid htis practice. It leads to the potential to create all sorts of bugs later. Just don't do it. To better avoid this, setting hte flag in teh settings (TODO: add a picture of that).

When decalring a vairable, there are two compoentns to it: variable name and variable type. Variable name is wholly your decision with only a couple of constraints. You are not allowed to duplicate the name of an internal comannd, and you shoudl go to some lenght ot avoid using the same name as an Excel object model name. Beware that naming a variable ahs certain concentions, but theese do not have any effect on the program execution. The main concern with names is that they will directly affect your ability to work with and maintain your code. Naming things is hard. Pick a strategy that works for you and your coworkers and get on it with it. There is no single answer here about how to name things.

The second part of the puzzle is to declare the type of the variable. This is THE core part of variables. When declaring a vairable, you are essentially deciding if the tpye should be the generic `Variant` or if you should actualyl declare a tpye. Note that there are times when you have ot use Variant, but in general, you should use the most speciifc tpye that is possible. These tpyes can either draw from VBA or from the Object Model, or from your own created types. When thinking of variable types, tehre are two major groups of types:

- Vlaue types = a number, string, or boolean
- Reference types = objects
