## things to change and check

This section will focus on the common properties that are checked and changed with these types of manipulations.

### properties of the `Range`

The common properties of the `Range` to work with include:

- `Value`
- `Text`
- `Formula`
- `Font`
- `Interior`
- `NumberFormat`

TODO: add some examples of working with these

### commonly used VBA functions

In addition to the properties of the `Range`, there are a handful of common VBA functions that come up when working with simple to moderate manipulations. These include:

- `Split` - split a string into an array based on a delimiter (the reverse of Join)
- `Join` - join an array into a string with a delimiter (this reverse of Split)
- `Asc` - determine the ASCII code for a character
- `Chr` - return a character for an ASCII code (the reverse of Asc)
- `InStr` - determine if a string is in another one (called Substring in other languages)
- `Left`, Mid, Right - grab parts of a string
- `Trim` - remove any whitespace from the start or end of a string
- `Len` - determine the length of a string
- `UCase`, LCase - used to force a string to upper or lower case
- `UBound`, LBound - determine the bounds of any array
- `WorksheetFunction` - get access to any Excel functions in VBA
- `IsNumeric`, IsEmpty - check if a number -- TODO: add the others here
- `CDbl`, CLng, CBool, CDate - convert a value of one type to another -- TODO: add any others
- `Replace` - replace one string in another
- `Application`.Index, Application.Match - these are the VBA versions of the Excel functions
- `Application`.Transpose - convert a 1D array from vertical to horizontal and back
- `Is` Nothing - check if a reference has been set
- `TypeName` - check the type of an object (useful if working with `Variant`)
- `RGB` - useful way to build colors from known red, green, and blue values
- `Count` - common way to get the size of a group, used often to resize an input/output or to check logic

TODO: search through bUTL for other common functions
