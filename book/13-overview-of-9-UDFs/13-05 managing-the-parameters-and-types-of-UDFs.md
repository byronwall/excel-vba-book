## managing the parameters and types of UDFs

This section will focus on a topic that is quite nuanced but can have a large impact on how reusable your UDF code is. The focus here is on how to specify the type of the parameters and possibly the return of the UDF.

The reason things get tricky is that Excel is able to feed a wide range of object types to a UDF depending on how it was called. The common types to see are:

- Range
- Array/Variant
- Double/Number
- String
- Date
- Error

The most common ways to call a UDF are

- Use a Range reference UDF(A1:B2)
- Use the result of some other operation UDF(5\*A2). This can result in different object
  - Array formula gives an array
  - Math might give a number
  - String formulas will give a string
  - IF or CHOOSE might allow for multiple options depending on the result

Given this wide range of choices, it's important to consider how you intend for you UDF to be called and what types of inputs you want to be able to handle. You can choose ot be as loose or as restrictive as you want on the parameter type, but this will have an impact on usage. If you go the loose route, you can call everything a Variant, but then you lose the utility of Intellisense as you are programming. If you go the strict route, you gain Intellisense, but might make your UDF fail on a simple case that it should be able to process.

As an example, let's say you've written a UDF that simple squares the number that it is fed. If you specify the parameter of this as a Range, your code will work fine with usages like UDF(A1), etc., but it will fail if someone sends in the result of math UDF(5\*A1). This is odd because assuming that A1 is a number, there is no reason that you cannot square the result of that. Instead however, you will get an error that the result of that math (which is a Double) cannot be converted to a Range and your code will error out. For a simple example like this, it makes the most sense to declare the parameter as a Variant and just rely on the Value being correct.

TODO: add code for that example

Things are fixed simple in that case, but it quickly becomes an issue when you want to handle different types of input. Maybe you are making a function that will concatenate an array of strings together. What happens when you only get a single string as a String instead of an Array containing Strings? Most likely, your code will fail in this instance, unless you've built int eh proper checks on the type. In this case, you will likely need to take a parameter of Variant and then do the checking to see how to handle it.

TODO: add an example of string concat code that works

The most common spot to see this sort of issue is when deciding whether to deal with a type of Range or Variant (to handle an array). It is nice to work directly with Ranges and avoid the Variant, but this will make your code weak against someone who wants to use an array formula to call your UDF. It typically does no take much work to process an Array, but it helps to design things from th start like that.

TODO: add before example of UDF using Range
TODO: add after example of that UDF using a Variant/Array instead of the Range
