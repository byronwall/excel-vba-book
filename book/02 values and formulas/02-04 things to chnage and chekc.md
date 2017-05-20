## things to change and check

This seciton will focus on the common properties that are checked and changed with these types of manipulations.

### properties of the `Range`

The common properties of the `Range` to work with include:

* Value
* Text
* Formula
* Font
* Interior
* NumberFormat

TODO: add some examples of working with these

### commonly used VBA functions

In addition to the properites of the `Range`, there are a handful of common VBA functojns that come up when workign with simple to moderate manipualatiojns.  These include:

Split - split a string into an array based on a delimeter (the reverse of Join)
Join - join an array into a string with a delimeter (this reverse of Split)
Asc - determine the ASCII code for a character
Chr - return a character for an ASCII code (the reverse of Asc)
InStr - determine if a string is in another one (called Substring in other languages)
Left, Mid, Right - grab parts of a string
Trim - remvoe any whitespace from the start or end of a string
Len - determine the length of a string
UCase, LCase - used to force a string to upper or lower case
UBound, LBound - determine the bounds of any array
WorksheetFuncton - get access to any Excel functiosn in VBA
IsNumeric, IsEmpty - chekc if a number -- TODO: add the others here
CDbl, CLng, CBool, CDate - convert a value of one type to another -- TODO: add any others
Replace - replace one string in another
Application.Index, Application.Match - these are the VBA versions of the Excel functions
Application.Transpose - convert a 1D array from vertical to horizontal and back
Is Nothing - check if a refernece has been set
TypeName - check the type of an object (useful if working with `Variant`)
RGB - useful way to build colors from known red, green, and blue values
Count - common way to get the size of a group, used often to resize an input/output or to check logic

TODO: searcht through bUTL for other common funcitons
