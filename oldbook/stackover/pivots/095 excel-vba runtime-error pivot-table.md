# SO item 095
Thanks for any input on this. I'm trying to make a simple pivot table that is taking data from sheet "5 Month Trending May 15" and putting it onto my Pivot Table sheet called "Errors By Criticality - Pivot".

When I try to set the pivot cache range with the create method it's returning a runtime error of type mismatch. I've checked the parameters and it looks like I'm setting it up correctly. I did try to specify the PivotTable version and was still getting the same error. My code is below.

I'm assuming it's something to do with the pvtCache variable or the way I'm setting it to the range but I can't figure any solutions out.

```
Sub PivotTableCode()

Dim pvtCache As PivotCache 
Dim pvt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem 
'Set the cache of the pivot table
Sheets("5 Month Trending May 15").Select
Set pvtCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, Range("A2:H38"))

'create the Pivot Table
Sheets("Errors by Criticality - Pivot").Select
Set pvt = ActiveSheet.PivotTables.Add(pvtCache, Range("AP2"), "MyPivotTable") 
End Sub

```

----

The [documentation for PivotCaches.Create](https://msdn.microsoft.com/en-us/library/office/ff839430.aspx) indicates

> The `SourceData` argument is required if `SourceType` isn't `xlExternal`. It can be a `Range` object (when `SourceType` is either `xlConsolidation` or `xlDatabase`) or an Excel Workbook Connection object (when `SourceType` is `xlExternal`).

Despite this, the macro recorder will always create a `String` here for the `SourceData`. [(It will even create a bad string if the `Sheet` has a space in the name).](http://stackoverflow.com/questions/30538131/run-time-error-5-invalid-procedure-call-or-argument/30540014#30540014)

Given the preference for the macro recorder, I often supply this as a `String` with the addresses.

I have been able before to supply a `Range` here so I am not certain what is specifically going on that prevents the `Range` usage in this case.

To use a `String`, your code would look like:

```
Set pvtCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, "'5 Month Trending May 15'!A2:H38")

```
