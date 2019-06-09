## calling a Sub or Function

When you are calling a Sub or Function, there are a couple of ways to do it. The preferred approach is to simply type the name of the Sub/Function aong with any required parameters. This will call the Sub. Another approach is to use `Call SubName` which is the same as `SubName`. This is an older approach that omse people prefer. It can sometimes be the case that Sub names are not particularly clear in the VBA and using `Call` has the effect of making it obvious that code flow is being directed into a Sub.

WHen calling a Function, you have the same approaches available. You can just use the Functiojn name or use call (TODO: is that right?).

ONe thing to be aware of with Functions is how to properly handle the return from the functiojn (assuming it actually returns something). This is where VBA gets a bit weird. THe rules here split on whether the Functiojn returns an Object or Value type.

For either type, you are reuiqred to call the Functiojn iwht parenthesesis. This signals to VBA: please retain and use the return of this Functiojn. For a reference type, you will need to use `Set` as required. For a value type, you will omit `Set`. See the code example below.

If you ever get the compile time error `Object reference not set` this means that you have not used a `Set` somewhere that is required. A good place to check are spots where you are using the return from a function. The same thing happens if you omit the parentheses. (TODO: is this right?)

```vb
Sub ExampleOfCallingCode()
    Dim rngReference as Range
    Set rngReference = someFunctionThatReturnsARange()

    Dim dblValue as Double
    dblValue = someFunctionThatReturnsADouble()

End Sub
```
