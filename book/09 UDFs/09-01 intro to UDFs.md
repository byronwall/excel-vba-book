## introduction to user defined functions (UDFs)

This chapter will focus on using VBA to create user defined functions (UDFs).  This area of VBA is so-named because it allows you to add functions that are callable from the spreadsheet.  Once you're familiar with VBA, you'll recognize that there is no differnece between a normal VBA Function and a UDF.  The only differnce is that a Functions "beocmes" a UDF once it is called from the spreadhseet. Having said that, UDFs are still increibly powerful and can be an incredible time saver when working wiht a spreadsheet.  The power of UDFs is that there are very few limitations to what you can do inside a UDF.  This means that you can do complicated tasks from a single function call in Excel.  Contrast this with the mess you get when doing complicated things with normal Excel functions.

This chapter will hit the major topics related to UDFs including:

* Debuggin them
* Working with variable types, especially parameters but including outputs
* Limitations of UDFs -- what you cannot do
* Limitations of UDFs in addins -- must have the addin
* Differnet applications of UDFs
    * Simple things - string functions, etc.
    * Complicated things
    * Duplicating Excel functiojnality in a simpler package
* Understanding volatility
* Understanding Ranges and how they relate to your function being called
* Hiding a VBA function from UDFs, and using the Option Private Module
* Building more powerful UDFs with ExcelDna

