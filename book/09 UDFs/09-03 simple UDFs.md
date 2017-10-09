## some simple UDFs

This section will focus on the "simple" UDFs.  It may sound silly, but there are a handful of surprisingly useful UDFs that are just a single line of code.  In general, these UDFs are used to return some information about the spreadsheet that you'd prefer Excel simply have a function for.  In later versions of Excel, some of these gaps have been filled (e.g. obtaining the formula for a cell) but sometimes these gaps still remain.  In addition to one-liners, there are a large number of simple UDFs that exist to replace a more complicated Excel formula.  These types of UDFs can be much easier to read and debug/test than a complicated array formula for example.  The final group of UDFs that comes up frequently is string processing.  Excel provides good functions for manipulating strings, but these cna be a complete pain without the use of helper columns.  A simple UDF can hold a variable which eliminates a lot of the need for helper columns via traditional formulas.

Before committing whole hog to UDFs being the best way to do things in Excel, it's important to remember that there are downsides to UDFs.  The most important is that if you want the UDF to live with the workbook (and not in an adding) then you are required to save the workbook as macro enabled.  This can be a deterrent to using them in certain environments.  THe other thing to remember is that UDFs can often be a crutch for not actually learning how to get the most out of Excel functions.  It can be easy and tempting (especially for an experienced programmer) to start blasting through a spreadsheet with UDFs instead of learning how to do something "the Excel way".  Depending on your work setting and who else will see your workbooks/code, this may be a bigger issue for some people.

### Common reasons for using a UDF

There are a number of consistent spots where I will use a UDF instead of fighting the Excel formulas.  Thee typically fall into a couple of categories:

Excel formulas can be quite complicated/repetitive if need to store a variable
Certain valuable pieces of information about the cell or a Range are not available via functions
Some things are just much easier to do with VBA than with Excel

### examples of simple UDFs

TODO: add some examples here of different types
