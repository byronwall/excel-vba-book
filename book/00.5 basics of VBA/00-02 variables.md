## declaring and setting variables

One of the core tasks when programming via VBA is working with variables.  Variables encompass a couple of different topics which makes sense since they are one of two core areas of VBA alongside control structures.  That is, your programming exists of two possible categories: variables and control structures.  Variables are made of the variables that you will need to delcare and set to make your program work while also including all of the various aspects of the Excel object model.  The object model is made of a significant number of variables (e.g. cell value for each cell) and a handful of Subs and Functions.  The variables that you declare and use will look very similar to the object that the Excel model is using.  There are also a large number of variables that you will create which exist to guide your own control structure or to encompass the algorithms that you need to execute.

To fully work with variables, it helps to split the topic into two areas: declaring variables and setting variables.  These two topics are quite simple when it comes time to type out the commands, but variable declaration adn setting is at the core of planning how a program will work.  The variable declaration will directly shape how the control strcutures will work.  The two go hand in hand and are equally important.

### Declaring Variables

Decalring vairables is a straight forward tasks.  VBA offers a simple command to declare a new variable: `Dim`.  Note that you can technically use a vriable before declaring it, but you should really avaoid htis practice.  It leads to the potential to create all sorts of bugs later.  Just don't do it.  To better avoid this, setting hte flag in teh settings (TODO: add a picture of that).

When decalring a vairable, there are two compoentns to it: variable name and variable type.  Variable name is wholly your decision with only a couple of constraints.  You are not allowed to duplicate the name of an internal comannd, and you shoudl go to some lenght ot avoid using the same name as an Excel object model name.  Beware that naming a variable ahs certain concentions, but theese do not have any effect on the program execution.  The main concern with names is that they will directly affect your ability to work with and maintain your code.  Naming things is hard.  Pick a strategy that works for you and your coworkers and get on it with it.  There is no single answer here about how to name things.

The second part of the puzzle is to declare the type of the variable.  This is THE core part of variables.  When declaring a vairable, you are essentially deciding if the tpye should be the generic `Variant` or if you should actualyl declare a tpye.  Note that there are times when you have ot use Variant, but in general, you should use the most speciifc tpye that is possible.  These tpyes can either draw from VBA or from the Object Model, or from your own created types.  When thinking of variable types, tehre are two major groups of types:

- Vlaue types = a number, string, or boolean
- Reference types = objects

### Setting variables

Setting a variable is quite straight forward.  The rule is: for reference tpyes, you must use `Set`, for value types, you must not.

The real task then is to determine whether or not you are workign wtih a refernece type.  The rule here is: if you are workign with an object, it is a refernece type.  If you are working with a value (number, stirng, bool), then you are working with a value.  Another approach, if you intend to use a `.` to call out some property of your variable, then it is a reference type and requires `Set`.  The one odd excepipon here is arrays: they are declared without using Set.
