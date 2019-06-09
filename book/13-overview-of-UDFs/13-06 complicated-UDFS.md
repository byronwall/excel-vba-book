## complicated UDFS

One of the great advantages of UDFs is that you give you full access to all of VBA while still executing within the spreadsheet. There are some limits to this power, but, in general, you able to do some very powerful stuff in the same interface that you nromally do a `SUM`. To take advantage of this power, you need to be aware that these things are possible and then consider taking a shot at it.

Some of the more complicate areas where you will want to write a UDF include:

- Using Range information from cells not related to the parameters
- Accessing the FileSystem

Related to the Range, you have full access to all of the Workbooks and Worksheets that are available in VBA. THis means that you can combine a large amount of data in VBA and then output it to a UDF return. Where this becomes useful if when you want to look at the metadata of a Range of Worksheet. Until Excel 2013, most of this information was simply not available without VBA. Post 2013, you are able to use the `CELL` function (TODO: is this right). Some of the more useful things here are to use the fomratting of a cell (e.g. return the background color) or the display value (i.e. `Range.Text`). These UDFs can be great for eitehr long term usage or for a quick throwaway to get infomraiton into the spreadsheet. When doing the latter, there is essentially no difference between using a UDF and running through the cells using a normal Sub. The main reason you might use a UDF is if the cells you want to target are not easy to identify in VBA.

Other possible UDFs allow you to access the file system and possibly return information frmo there. ONe example would be to return the size of a file in KB given a file name. Really, you could go get any information you want. Again, this type of UDF can be easily done as a UDF or just as a Sub that runs through a Range as an input.

When considering whether or not to use a UDF or a Sub, consider the following:

- A UDF will update automatically when the parameters to it change (or always if marked as Volatile). This is the key differnentiator.
- A Sub can run without embedding itself into a spreadsheet. This is key if you need to save the spreadsheet with your information without a link to your code. This is a moot point if your UDF lives in an XLSM file but starts to matter for an addin. You can also do a copy/paste values if you want to remove the UDF.

TODO: consider adding more here or refining this section
