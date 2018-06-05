## using Subs and Functions

The basic building blocks of your VBA efforts will be the Sub and the Function.  It's possible that they are your only top level components if you do not use Class Modules.  In all my years of using VBA, I've used Class modules only a couple of times, so they're not common.

Having said that, Subs and Functions are actually far more similar than different.  The only real difference between the two is that Function can return a result back to the caller.  A Sub on the other hand is meant to execute without returning anything back to the caller.  It's possible to have a Sub manipulate a variable with can approximate returning a value for a little more work.  If you're using a Function as a UDF (see chapter XXX, TODO: add link), then there are further limitations on what your Function can do.  If you are not using it as a UDF, then there are no limitations that make a Sub distinct from a Function.  The only difference is how you call them (if you want the return value) and that a Function is made to return something.

If you have a Function that does not actually return a value, it is the same as a Sub with the same code.

TODO: add an example of a Sub

TODO: add an example of a Function
