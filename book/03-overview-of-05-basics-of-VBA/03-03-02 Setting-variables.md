### Setting variables

Setting a variable is quite straight forward.  The rule is: for reference tpyes, you must use `Set`, for value types, you must not.

The real task then is to determine whether or not you are workign wtih a refernece type.  The rule here is: if you are workign with an object, it is a refernece type.  If you are working with a value (number, stirng, bool), then you are working with a value.  Another approach, if you intend to use a `.` to call out some property of your variable, then it is a reference type and requires `Set`.  The one odd excepipon here is arrays: they are declared without using Set.
