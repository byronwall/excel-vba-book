# SO item 019
So i am trying to dynamically insert data into an Excel table from other closed workbooks.

**i got everything working just fine, except one small annoying thing.**

i have a formula as follows:

> ='H:\dev...[some book name.xlsm]Main'!C1

the formula above works fine. what i need is to insert this exact same formula into a table in the sheet for several rows.

it should look like that in one column:

> ='H:\dev...[some book name.xlsm]Main'!C1
> ='H:\dev...[some book name.xlsm]Main'!C1
> ='H:\dev...[some book name.xlsm]Main'!C1

what excel does, it automatically changes all the cell references to be incremental, like this:

> ='H:\dev...[some book name.xlsm]Main'!**C1**
> ='H:\dev...[some book name.xlsm]Main'!**C2**
> ='H:\dev...[some book name.xlsm]Main'!**C3**

i insert the the formulas as a string into an array, and then paste it into the table using this code:

```
Set lstObj = Sheets(1).ListObjects(1)
Set rngLstObj = lstObj.Range
With rngLstObj.Offset(1, 0).Resize(rngLstObj.Rows.Count - 1,rngLstObj.Columns.Count)
    .Formula = RevList  
End With

```

in the code above, `RevList` is a 2 dimentional array.

i tried setting it to `.Formula`, `.Value`, in both cases excel changes the cell references to be incremental.

i tried disabling calculation

```
ThisWorkbook.Sheets(1).EnableCalculation = False

```

still same.

how do i stop this behavior from VBA side?

----

If you start with

```
='H:\dev...[some book name.xlsm]Main'!$C$1

```

it will force the absolute reference wherever it is copied.

The added dollar signs prevent the range from changing.
