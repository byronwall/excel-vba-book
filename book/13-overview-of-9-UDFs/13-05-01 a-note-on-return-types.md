### a note on return types

THe same thing can happen on the return side of the equation, but it is typically less of a problem.  The main issues on the return side are returning arrays and dealing with Strings.  If you want your UDF to work as an array formula, you can simply return an array and it will work.  If that array is only a single cell, then it will look the same as a non-array formula.

Another issue is when working with Strings.  If you return a string from a UDF, it will be formatted as Text instead of General.  TODO: is that true?  THis can have intended consequences as Excel tends to treat Text differently when it is then sent to other functions.  THe most common example is that a number stored as text will not be available for normal math operations.

You can avoid this by returning Variant but it can become an issue when yuo want a Function to work as a UDF and as a normal VBA Function.  You might have a good reason to use a specific return type on the VBA side of things, but then Excel may not handle that the way you want (if using a String).  Or, going the other way, you may have a UDF that works great because Excel can treat a single entry array as a single cell, but that becomes complicated when you call the UDF from another VBA location and then have to deal with a single number versus an array.
