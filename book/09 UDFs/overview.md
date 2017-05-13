# overview of 9 UDFs

This chapter could focus on a couple aspects of UDFs.  High level topics:

* Using them to return simple info that is hard to get elsewise (e.g. Range.Formula)
* Using them to hide complicated logic that could be done in a formula but would be a mess
* Using them to do things that are not possible otherwise

UDFs are a great way to extend Excel with some common features

Could include some examples of where this has been done in bUTL:

* String processing is much easier with UDFs instead of formulas (concatenation)
* Doing logic that might otherwise require an array formula
* UDFs are a great way to simplify formulas for conditional formatting
* UDFs are a great addition to a personal addin where the funcitonality is available without copying/changing formulas

Some technical points to hit:

* THe pitfalls of using Ranges outside of the ones referred to
* Making a funciton Volatile and what that means
