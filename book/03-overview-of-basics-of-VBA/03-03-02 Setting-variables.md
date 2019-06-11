### Setting variables

Setting a variable is quite straight forward. The rule is: for reference types, you must use `Set`, for value types, you must not.

The real task then is to determine whether or not you are working with a reference type. The rule here is: if you are working with an object, it is a reference type. If you are working with a value (number, string, bool), then you are working with a value. Another approach, if you intend to use a `.` to call out some property of your variable, then it is a reference type and requires `Set`. The one odd exception here is arrays: they are declared without using Set.
