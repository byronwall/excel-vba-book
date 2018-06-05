### debugging UDFs is different

Most folks are familiar with the "Evaluate Function" feature of Excel which will help you walk through a function's evaluation in the order that Excel evaluates things.  This can be incredibly helpful for array formulas where it's not always obvious the order Excel will do things in.  Your UDF will also be evaluated in that feature, but it will not step through the logic of your UDF.  This might seem obvious, but it's worth mentioning.  IF you want to debug the logic of your addin, you need to set a breakpoint and actually debug the code.  See the later section on this for the details.

TODO: add link to that section
