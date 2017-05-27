## managing the parameters and types of UDFs

This section will focus on a topic that is quite nuanced but can have a large impact on how reusable your UDF code is.  The focus here is on how to specify the type of the parameters and possibly the return fo the UDF.

The reaosn things get tricky is that Excel is able to feed a wide range of object types to a UDF depending on how it was called.  The common types to see are:

* Range
* Array/Variant
* Double/Number
* String
* Date
* Error

The most common ways to call a UDF are

* Use a Range refernece UDF(A1:B2)
* Use the result of some otehr operation UDF(5*A2).  This can result in different object
    * Array formula gives an array
    * Math might give a number
    * String fomrulas will give a string
    * IF or CHOOSE might allow for multiple options depending on the result

Given this wide range of choices, it's important to conisder how you intend for you UDF to be called and what types of inputs you want to be able to handle.  You can choose ot be as loose or as restrictive as you want on the parameter type, but this will have an impact on usage.  If you go the loose route, you can call everything a Variant, but then you lose the utility of Intellisense as you are programming.  If you go the strict route, you gain INtellisense, but might mkae your UDF fail on a simple case that it should be able to process.

As an example, let's say you've written a UDF that simple squares the number that it is fed.  If you sepcify the parmaeter of this as a Range, oyur code will work fine with usages like UDF(A1), etc., but it will fail if someone sends in the result of math UDF(5*A1).  This is odd because assuming htat A1 is a nujmber, there is no reason that you cannot square the result of that.  Instead however, you will get an error that the result of that math (which is a Double) cannot be converted to a Range adn your code will error out.  For a simple example like this, it makes the most sense to declare the parameter as a Variant and just rely on teh Value being correct.

TODO: add code for that example

Things are fixed simple in that case, but it quickly beocmes an issue when you want to handle differnet types of input.  Maybe you are making a function that will concatenate an array of strings together.  What happens when you only get a single string as a Stirng instead of an Array containig Strings?  Most liekly, your code will fail in this instance, unless you've built int eh proper checks on teh type. In this case, you will likely need to take a parameter of Variant and hten do the checking to see how to handle it.

TODO: add an example of string concat code that works

The msot common spot to see this osrt of issue is when deciding whetehr to deal with a type of Range or Variatn (to hadnle an array).  It is nice to work directly with Ranges and avoid the Variant, but this will make your code weak against someone who wants to use an array formula to call your UDF.  IT typically does no take much work to process an Array, but it helps to desing things from th start like that.

TODO: add before example of UDF using Range
TODO: add after example of that UDF using a Variant/Array instead of the Range

### a note on return types

THe same thing can happen on teh retunr side of the equation, but it is typically less of a problem.  The main isseus on the reutnr side are returnign arrays adn dealing with Strings.  If you want your UDF to work as an array formula, you can simply return an array and it will work.  If that array is only a single cell, then it will look the same as a non-array formula.

Another issue is when working with Strings.  If you retunr a string from a UDF, it will be formatted as Text instead of General.  TODO: is that true?  THis can have intended consequences as Excel tends to treat Text differently when it is then sent to other functions.  THe most common example is taht a number stored as text will not be avialable for normal math operations.

You can avoid this by returning Variant but it can beocme an issue when yuo want a Funciton to work as a UDF and as a normla VBA Function.  You might have a good reason to use a specific return type on teh VBA side of things, but then Excel may not hanlde that the way you want (if using a String).  Or, going the other way, you may have a UDF that works great because Excel can treat a singel entry arrya as a single cell, but that becomes complicated when you call the UDF from another VBA location adn then have to deal with a singel nujmber versus an array.
