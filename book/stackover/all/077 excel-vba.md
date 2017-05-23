# SO item 077
I am having a heck of a time having code from one workbook ("W1") open another workbook ("W2") and perform deletions in W2\. When I run it, it selects the ranges in W2 but will not delete the selection. I figured out I must be explict with naming W2 for the deletions, but I'm getting lost with it. Any help would be very appreciated.

My code is as follows:

```
Sub Clear_FM_Contents()

Dim f As FileDialog
Dim varfile As Variant
Dim path As Variant

'Prompt the user to select the Excel File to Import
Set f = Application.FileDialog(msoFileDialogFilePicker)

'Error handling with file selector
If f.Show = False Then
    MsgBox "You clicked Cancel in the file dialog box."
    End
End If

'Set the path of the User selected file
For Each varfile In f.SelectedItems
    path = varfile
Next

'Create the Excel object
Dim xlApp As Excel.Application
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True

'Open the selected Excel file
xlApp.Workbooks.Open path, True, False

'Clear all Template Inputs
With xlApp.ActiveWorkbook
    .Sheets("Mrkt Data").Select
    With xlApp.ActiveWorkbook.ActiveSheet
        .Range("E36,E5,E7,E10,E13,E17,E39,J7:J11,J13:J20,J24:J29,J31:J33,J35,O21").Select
        .Range("O21").Activate
        xlApp.ActiveWorkbook.ActiveSheet.Selection.ClearContents
    End With
    'Close the Excel File
    ActiveWorkbook.Save
    ActiveWorkbook.Close  
End With
'Close Excel
xlApp.Quit
'Eliminate the xl app object from memory
Set xlApp = Nothing

MsgBox "Model Inputs Cleared"
End Sub

```

----

If `Select` works, then just replace it with `ClearContents` instead and delete the 2 lines under it that try to `Activate` and `Clear`. I suspect the issue is `Activate` and that `O21` is being cleared but not the others. You could also just delete the `Activate` line but I am trying to clean up an instance of `Select` that is not needed.

```
With xlApp.ActiveWorkbook.ActiveSheet
    .Range("E36,E5,E7,E10,E13,E17,E39,J7:J11,J13:J20,J24:J29,J31:J33,J35,O21").ClearContents
End With

```

instead of

```
With xlApp.ActiveWorkbook.ActiveSheet
    .Range("E36,E5,E7,E10,E13,E17,E39,J7:J11,J13:J20,J24:J29,J31:J33,J35,O21").Select
    .Range("O21").Activate
    xlApp.ActiveWorkbook.ActiveSheet.Selection.ClearContents
End With

```

You could also compress the `With` and do it all on one line, but that is too wide for here so I didn't do it.

You should also take a look at [this post about not using Select.](http://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba-macros/10717999#10717999)
