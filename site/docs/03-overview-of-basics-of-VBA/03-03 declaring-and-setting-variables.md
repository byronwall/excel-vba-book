## Declaring and Setting Variables

One of the core tasks when programming via VBA is working with variables. Variables are used to reference the Excel object model and to guide control structures. Within the Excel object model, the objects hold variables which point to other objects. Working with these objects is critical to using VBA. You will need to understand variables to do that.

This section is split into two areas: declaring variables and setting variables. The code for these two topics is simple. The complexity comes in planning out the best structure for managing variables. The variable declaration will directly shape how the control structures will work.

### Declaring Variables

Declaring variables is straight forward. VBA offers a simple command to declare a new variable: `Dim`.

When declaring a variable, there are two components to it: variable name and variable type. Variable names are your choice with some constraints. You are not allowed to duplicate the name of an internal command, and you should go to some length to avoid using the same name as an Excel object model name. Beware that naming a variable has certain conventions, but these do not have any effect on the program execution. The main concern with names is that they will directly affect your ability to work with and maintain your code. Naming things is hard. Pick a strategy that works for you and your coworkers and get on it with it. There is no single answer here about how to name things.

The second part of the puzzle is to declare the type of the variable. This is THE core part of variables. When declaring a variable, you decide if the type should be the generic `Variant` or if you need a more specific type. There are times when you have to use Variant, but you should aim to use the most specific type that is possible. These types draw from VBA, from the Excel Object Model, or from your own created types. When thinking of variable types, there are two major groups of types:

- Value types = a number, string, or boolean
- Reference types = objects

TODO: find better place:

Note that you can technically use a variable before declaring it, but you should really avoid this practice. It leads to the potential to create all sorts of bugs later. Just don't do it. To better avoid this, setting the flag in the settings (TODO: add a picture of that).

TODO: add code sample for declaring a variable (show an object, primitive, and array)

### Setting variables

Setting a variable is straight forward. The rule is: **for reference types, you must use `Set`; for value types, you must not.**

The real problem then is to determine whether or not you are working with a reference type. The rule is: if you are working with an object, it is a reference type. If you are working with a value (number, string, boolean), then you have a value type. Another approach, if you intend to use a `.` to call out some property of your variable, then it is a reference type. The exception here is arrays: they are set without using Set.

TODO: add code sample showing variable setting

### Using Variables

It seems somewhat obvious that you would want to use a variable after declaring and setting it. This is generally always the case (why else would you create the variable). To that end, there are a pair of ways to use variables depending on whether it is a reference or value type. Value types are easier since you can only do 1 thing with them: use them in an expression. This feels and usually looks like mathematical formulas. The more complicated example comes with reference types where the variable stores a reference to another object. These variables have the ability to access either a property of the type or the default `Value` of the type. The distinctions between reference and value types can become confusing with the Excel Object Model since so many properties of objects reduce to value types. An example is the value of a `Range` which will hold some number or string or Error depending on what the cell contains.

When accessing a property of the object, you use the `.` to access a property by name. In this way, you can chain together a series of commands accessing the properties of objects. It is often the case that the property is itself another object which makes it possible to use another `.` to keep going. If you are using the VBE and properly declaring your variables, the VBE will work to provide helpful suggestions of what may be possible to use next (this is called Intellisense). The one pitfall to Intellisense is when the return from a given property can be Variant or a combination of possible results. When this happens, Intellisense will not offer any suggestions and you are left guessing whether or not the command exists. This is where it can be quite helpful to do one of two things:

TODO: create a demo of these bullets

- Create a new variable with the type that you know the object will have and Set that reference before using it. This "cheats" and tells Intellisense exactly what you expect to exist.
- Read through the documentation and gain an understanding of what types are possible and just use them. There is no rule that the type must be suggested by Intellisense for it to be valid.

In general, I take a combination of those two approaches often. If I expect to use the variable a number of times, I will go with the new variable route to avoid guessing properties later. If I only need the variable once or am copying code from somewhere else (and know it works), I will just go with the code as is without Intellisense. The one upside of creating new variables is that it forces you to be more explicit with your declarations. It also clearly shows your intent to other developers that may see your code later.

### Value Default

I mentioned it above, but it is worth digging into the default `Value` property a little more. This can be a source of confusion because very often, you will accidentally use the name of a variable without calling for a property. In other programming languages, this will result in a compile time or runtime error. In VBA, your code will run and even worse will return something from the object that may not be what you want. When this happens, it can be incredibly difficult to track down the source of the error. To avoid this, you could never use the variable name as a shortcut to the `.Value` property. In practice this is a pain to manage and I will often mix and match whether or not Value is called. Sometimes, I am tired of typing out Value and just let the default work. Other times, I am being very diligent about calling everything explicitly to avoid some unforeseen error later. You will find that this comes down to your own preference and the preferences of others working on your code.
