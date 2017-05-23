# SO item 035
I have an Excel spreadsheet that has add-in functions that compute on data given in a fixed range. All the worksheets work properly with pressing of F9 or Shift-F9 keys.

I am writing a for loop with VBA. It copies a range from one worksheet which contains all the data to another worksheet. Then calculates manually one worksheet at a time, with pauses, and even compute twice so as to ensure the execution of each worksheet. If I manually step through the VBA code line by line in the debug mode through the whole for loops, everything works. However, if I press F5 and run the whole VBA for loop at full speed. The code produces the same result for many (not all) of the consecutive iterations of the for loop, while I know the results should all be different. My guess is that the workseets become stale and are not replaced with new data.

The following is my VBA code. I would really appreciate if someone would take a look at it and help to resolve this problem. The main part of the code that is problematic is "sht.Calculate".

```
Option Base 1
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Function runCalibrate()
Dim xl As Workbook
Set xl = ThisWorkbook
nRowPerBlock = 15
yrs = Array("2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015")

Application.Calculation = xlCalculationManual

strikeSentinel = Array(1, 13)
maturitySentinel = Array(1, 10)

nRow = maturitySentinel(2) - maturitySentinel(1) + 1
nCol = strikeSentinel(2) - strikeSentinel(1) + 1
For j = 1 To UBound(yrs)
    Set a = Worksheets(yrs(j)).Cells(3, 2)
    For i = 0 To 11
        b = a.Offset(1 + nRowPerBlock * i, 1).Resize(nRow, nCol)
        xl.Sheets("SPX").Range("p4:ab13") = b

        'refresh required worksheets
        Call RefreshSheetNX(xl.Sheets("TODAY"))
        Call RefreshSheetNX(xl.Sheets("USD"))
        Call RefreshSheetNX(xl.Sheets("SPX"))
        Call RefreshSheetNX(xl.Sheets("EQ Model"))
        Call RefreshSheetNX(xl.Sheets("EQ Calibration"))

        c = xl.Sheets("EQ Calibration").Range("c18:c32")
        xl.Sheets("calibration time series").Cells(38, 2 + 12 * (j - 1) + i).Resize(15, 1) = c
    Next i
Next j

End Function

Function RefreshSheetNX(sht As Worksheet)
    On Error Resume Next
    sht.Cells.Dirty
    sht.Calculate
    Call Sleep(1000)
    sht.Calculate
    Call Sleep(1000)
End Function

```

* * *

As Byron suggested, I use 'Application.CalculateFullRebuild' as in the following code. But it is not reacting at all when I stepped through 'Application.CalculateFullRebuild'.

```
Option Base 1
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Function runCalibrate()
Dim xl As Workbook
Set xl = ThisWorkbook
nRowPerBlock = 15
yrs = Array("2005", "2006", "2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015")

Application.Calculation = xlCalculationManual

strikeSentinel = Array(1, 13)
maturitySentinel = Array(1, 10)

nRow = maturitySentinel(2) - maturitySentinel(1) + 1
nCol = strikeSentinel(2) - strikeSentinel(1) + 1
For j = 1 To UBound(yrs)
    Set a = Worksheets(yrs(j)).Cells(3, 2)
    For i = 0 To 11
        b = a.Offset(1 + nRowPerBlock * i, 1).Resize(nRow, nCol)
        xl.Sheets("SPX").Range("p4:ab13") = b

        'refresh required worksheets
'        Call RefreshSheetNX(xl.Sheets("TODAY"))
'        Call RefreshSheetNX(xl.Sheets("USD"))
'        Call RefreshSheetNX(xl.Sheets("SPX"))
'        Call RefreshSheetNX(xl.Sheets("EQ Model"))
'        Call RefreshSheetNX(xl.Sheets("EQ Calibration"))
        Application.CalculateFullRebuild            

        c = xl.Sheets("EQ Calibration").Range("c18:c32")
        xl.Sheets("calibration time series").Cells(38, 2 + 12 * (j - 1) + i).Resize(15, 1) = c
    Next i
Next j

End Function

```

----

Consider the use of `Application.CalculateFullRebuild` instead of multiple calls to `Calculate` for each sheet. It will force a recalc of everything and rebuild dependencies (as the name implies).

Depending on how formulas are related between sheets, you might not be triggering the right order of calculations processing each sheet one at a time. The full rebuild approach is the "nuclear option" for refreshing a spreadsheet. Specifically, calling `sht.Cells.Dirty` and then `sht.Calculate` may only force a refresh of the current sheet without consideration for changes on other sheets.
