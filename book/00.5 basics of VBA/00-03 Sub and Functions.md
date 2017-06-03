## using Subs and Functions

The basic building blocks of your VBA efforts will be the Sub and the Function.  It's possible that they are your only top level components if you do not need to use your own Class Modules.  In all my years of using VBA, I've used Class modules only a couple of times, so they're not common.

Having said that, Subs and Functions are actually far more similar than different.  The only real difference between the two is that Function can return a result from its call back to the caller.  A Sub on the other hand is meant to execute without returning anything back to the caller.  It's possible to have a Sub manipulate a variable with can approximate returning a value for a little more work.  If you're using a Function as a UDF (see chapter XXX, TODO: add link), then there are further limitations on what your Function can do.  If you are not using it as a UDF, then there are no limitations that make a Sub distinct from a Function.  The only difference is how you call them (if you want the return value) and that a Function is made to return something.

If you have a Function that does not actually return a value, it is the same as a Sub with the same code.

TODO: add an example of a Sub
TODO: add an example of a Function

## declaring the parameters (Subs and Functions)

TODO: add content

### declaring an Optional parameter

TODO: add content

## calling a Sub or Function

TODO: add content

## declaring the return type (Function only)

For a Function, the only extra step is to declare the return type of the Function.  This is done after the normal parameters, with an extra `as Type` where `Type` is the actual type that you want to return.  Note that this type must be compatible with all possible Types that you could return.  Sometimes this means that you need to return a Variant in order to have all possible return Types available to you.  There are times where this makes sense (and a large part of the Excel object model does this), but note that using Variant will make it hard to use Intellisense to figure out what your VBA is capable of doing.

TODO: is this a Variant by default?
TODO: give some examples of Function returns

### returning from a Function

If you want to take advantage of a Function, you need to return a value from your Function.  This returned value can then be consumed by the caller (or not).  To return a value from a Function, you simply use the Function name as a variable and set its value appropriate.  If the return type is an object or reference type, then you need to use Set to return teh object.  If it is a value type instead, you can simple set the return with an equal statement like any other value type.  Once you have made the return statement, you can call Exit Function to break out of the Function.

For the caller, there are two things to keep in mind when using Functions.  The first is that you must call the Function with parentheses in order to access the return value.  The corollary of this is that if you call a Function with parentheses, you must use that return value to set the value of a variable.  You will get an error if you do not do this correctly.  Note that if you do not want the return value for some reason, you can avoid using parentheses in teh same way you call a Sub.  The second part is that you must call Set if the variable is an object/reference and not a value.

TODO: give an example of the return type and returning
