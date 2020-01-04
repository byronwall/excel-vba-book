### Setting variables

Setting a variable is straight forward. The rule is: **for reference types, you must use `Set`; for value types, you must not.**

The real problem then is to determine whether or not you are working with a reference type. The rule is: if you are working with an object, it is a reference type. If you are working with a value (number, string, boolean), then you have a value type. Another approach, if you intend to use a `.` to call out some property of your variable, then it is a reference type. The exception here is arrays: they are set without using Set.

TODO: add code sample showing variable setting
