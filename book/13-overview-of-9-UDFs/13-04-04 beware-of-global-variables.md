### beware of global variables

VBA allows you to declare a variable outside of any Sub or Function definition.  These are typically called global variables because they can be accessed from any code.  This means that you can create some variables in a Sub and then use them in subsequent UDF calls.  A good example is loading up a database of information and then using that information inside the UDF.  This cna be nice because then you do not have to load the data every time you call the UDF.  I've used this effectively when doing unit conversions with UDFs.

The downside to this approach is that it seems to be relatively easy to corrupt those global variables if you have errors while the UDF runs.  I've had it happen where that loaded database becomes corrupted somehow and then all of the dependent cells start to fail when their UDF is called.  This type of error can be quite difficult to track down because it may not be obvious why the variable was corrupted.
