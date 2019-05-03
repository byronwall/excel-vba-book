## introduction to user defined functions (UDFs)

This chapter will focus on using VBA to create user defined functions (UDFs). This area of VBA is so-named because it allows you to add functions that are callable from the spreadsheet. Once you're familiar with VBA, you'll recognize that there is no difference between a normal VBA Function and a UDF. The only difference is that a Functions "becomes" a UDF once it is called from the spreadsheet. Having said that, UDFs are still incredibly powerful and can be an incredible time saver when working with a spreadsheet. The power of UDFs is that there are very few limitations to what you can do inside a UDF. This means that you can do complicated tasks from a single function call in Excel. Contrast this with the mess you get when doing complicated things with normal Excel functions.

This chapter will hit the major topics related to UDFs including:

- Debugging them
- Working with variable types, especially parameters but including outputs
- Limitations of UDFs -- what you cannot do
- Limitations of UDFs in addins -- must have the addin
- Different applications of UDFs
  - Simple things - string functions, etc.
  - Complicated things
  - Duplicating Excel functionality in a simpler package
- Understanding volatility
- Understanding Ranges and how they relate to your function being called
- Hiding a VBA function from UDFs, and using the Option Private Module
- Building more powerful UDFs with ExcelDna
