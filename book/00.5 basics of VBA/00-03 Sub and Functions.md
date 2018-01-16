## using Subs and Functions

The basic building blocks of your VBA efforts will be the Sub and the Function.  It's possible that they are your only top level components if you do not use Class Modules.  In all my years of using VBA, I've used Class modules only a couple of times, so they're not common.

Having said that, Subs and Functions are actually far more similar than different.  The only real difference between the two is that Function can return a result back to the caller.  A Sub on the other hand is meant to execute without returning anything back to the caller.  It's possible to have a Sub manipulate a variable with can approximate returning a value for a little more work.  If you're using a Function as a UDF (see chapter XXX, TODO: add link), then there are further limitations on what your Function can do.  If you are not using it as a UDF, then there are no limitations that make a Sub distinct from a Function.  The only difference is how you call them (if you want the return value) and that a Function is made to return something.

If you have a Function that does not actually return a value, it is the same as a Sub with the same code.

TODO: add an example of a Sub

TODO: add an example of a Function

## declaring the parameters (Subs and Functions)

When creating a new Sub or Function you are able to determine the inputs to your new creation.  There are a handful of ways of handling the inputs:

* Put the inputs into the parameters of the Sub/Function and allow the caller to provide them
* Use knowledge of the spreadsheet to determine the inputs (or prompt the user for an input)

The main split here is: do you require the person typing the VBA to give you the inputs?  Or, do you use some other approach like asking the user or just pulling the inputs from the spreadsheet.

The most common approach is to pull the inputs out of the spreadsheet.  This seems counterintuive, but if you consider that the vast majority of VBA code is purpsoes wirtten for a single use, then it stands to reason that code will not be built on a large nujmber of Subs/Functiojns accepting parameters.  The reason for this is that generally someone writes VBA to handle *their* spreadhseet and so the VBA just refelcts that spreadsheet.  This works great for indivudal cases but can become a burden when building larger workflows.  The main thign to consider for lager workflows is that as the complexity grows, tehre will be a large amount og code that is called multiple times or could be called separately from the main workflow.  When thi sis the case, you are often served by pulling that code out into its own Sub/Functiojn wiiht aprameters.

To create a Sub or Function with parmaeters, you simply add them to the defintion line:

```vb
Sub WithSomeName(firstParameter as String)

End Sub
```

This approach is very simple.  You give the parameter a name and a type delcaration.  This is very nice becasue it nearly exaclty matches the `Dim` statement with a Sub.  That correspondence makes it very easy to start with an internally declared variable and then upgrade it to parmaeter.  You can also go the other way: take a parmaeter and inline it into the Sub with some default or determined value.  This is less common.

Once the parameter has been given a name and a type, you can simply use it within the Sub like any other variable.  In this regard, your code will look teh exact same.  IF you are the person typing the VBA to use this Sub, then you will have to provide an appropriate variable as the parameter to make it all work.

### declaring an Optional parameter

The one additiojnal thign to consider is that of `Optional` parameters.  An optional parameter is one who is not strictly requried.  In liue of a value, you can either leave the parmaete rmissing or provide a default value.  In iether case, you can use the VBA specific function `IsMissing()` to determine if the aprmaeter was entered.  An Optiojnal parameter can be a very nice fearture when you are trying to determine whether or not to make a Sub take parameters or just use defaults.  You can provide the defaults in the parmaeter declaaration adn then allow the user (person typing the VBA) to override them if needed. This is a very common approach when writing library type code; provide snesible defualts that can be overwritten.

## calling a Sub or Function

When you are calling a Sub or Funciton, there are a couple of ways to do it. The preferred approach is to simply type the name of the Sub/Function aong with any required parameters.  This will call the Sub.  Another approach is to use `Call SubName` which is the same as `SubName`.  This is an older approach that omse people prefer.  It can sometimes be the case that Sub names are not particularly clear in the VBA and using `Call` has the effect of making it obvious that code flow is being directed into a Sub.

WHen calling a Funciton, you have the same approaches available.  You can just use the Functiojn name or use call (TODO: is that right?).

ONe thing to be aware of with Functiojns is how to properly handle the return from teh functiojn (assuming it actually returns something).  This is where VBA gets a bit weird.  THe rules here split on whether the Functiojn returns an Object or Value type.

For either type, you are reuiqred to call the Functiojn iwht parenthesesis.  This signals to VBA: please retain and use the return of this Functiojn.  For a reference tpye, you will need to use `Set` as required.  For a value type, you will omit `Set`.  See the code example below.

If you ever get the compile time error `Object reference not set` this means that you have not used a `Set` somewhere that is required.  A good place to check are spots where you are usign the retunr from a function.  The same thing happens if you omit the parentheses.  (TODO: is this right?)

```vb
Sub ExampleOfCallingCode()
    Dim rngReference as Range
    Set rngReference = someFunctionThatReturnsARange()

    Dim dblValue as Double
    dblValue = someFunctionThatReturnsADouble()

End Sub
```

## declaring the return type (Function only)

For a Function, the only extra step is to declare the return type of the Function.  This is done after the normal parameters, with an extra `as Type` where `Type` is the actual type that you want to return.  Note that this type must be compatible with all possible Types that you could return.  Sometimes this means that you need to return a Variant in order to have all possible return Types available to you.  There are times where this makes sense (and a large part of the Excel object model does this), but note that using Variant will make it hard to use Intellisense to figure out what your VBA is capable of doing.

TODO: is this a Variant by default?

TODO: give some examples of Function returns (or link to examples of them)

### returning from a Function

If you want to take advantage of a Function, you need to return a value from your Function.  This returned value can then be consumed by the caller (or not).  To return a value from a Function, you simply use the Function name as a variable and set its value appropriate.  If the return type is an object or reference type, then you need to use Set to return the object.  If it is a value type instead, you can simply set the return with an equal statement like any other value type.  Once you have made the return statement, you can call Exit Function to break out of the Function.

For the caller, there are two things to keep in mind when using Functions.  The first is that you must call the Function with parentheses in order to access the return value.  The corollary of this is that if you call a Function with parentheses, you must use that return value to set the value of a variable.  You will get an error if you do not do this correctly.  Note that if you do not want the return value for some reason, you can avoid using parentheses in the same way you call a Sub.  The second part is that you must call Set if the variable is an object/reference and not a value.

TODO: give an example of the return type and returning
