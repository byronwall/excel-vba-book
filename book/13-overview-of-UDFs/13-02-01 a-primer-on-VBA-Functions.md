### a primer on VBA Functions

Check the start of this book for a proper review of VBA Functions. The key points when using a function to execute as a UDF are:

TODO: link to section

- Function needs to be declared as Public
- Function needs to have a return type that can be processed in a cell (has a Value)
- Function needs to return something
- Function needs to be created in a code Module (not in a Worksheet or Workbook object)

Once you've met these criteria, you will be off and running. Typically a UDF will not work for one of those three reasons above. In particular, I regularly forget to declare the function Public and put it into a module. It's typically easier to remember to set the return type, but it is possible to forget ot actually return something from the function

The best indicator of whether or not these steps have been followed is to type your UDF into a spreadsheet and see if it is recognized. Excel does a very good job of identifying valid functions and offering them in the autocomplete.

Tip: Sometimes it is difficult ot remember the parameters that a UDF takes. You can either use the function input helper (TODO: add details about that) ro you can use the shortcut CTRl+SHIFT+A which will populate the names of the parameters into the UDF. Note that these are unlikely to be valid inputs to the function, so you will actually need to update the parameters. If you use descriptive names for the parameters (which you should!), this is a very helpful shortcut.

TODO: add an example of a very simple UDF here
