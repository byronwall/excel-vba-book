# SO item 082
I have reviewed many of the various examples of how to separate out tabs in workbooks into separate workbooks and have found one which works really well for what I need it:

*   creates a date/time stamped folder off the current folder from which the main file is stored;

*   it copies all visible sheets into separate worksheets naming each file the same as the tab name;

The issue for me is the size of the each new file is running at circa 8-10Mb as I suspect the underlying pivots and data files are carried over. What I need is to have separate files with just the values and formatting (plus column widths ideally).

I have looked at the code and it seems to use sh.copy but I cannot see where it decided to paste - hence cannot see how to close this down with paste value etc . It might be the syntax of sh.copy is just to follow this with the new destwb and this implies paste - but my knowledge of VBA is not up to altering this. The current code which I am looking to amend is:

```
Sub Copy_Every_Sheet_To_New_Workbook()
'Working in 97-2013
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim sh As Worksheet
    Dim DateString As String
    Dim FolderName As String

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    'Copy every sheet from the workbook with this macro
    Set Sourcewb = ThisWorkbook

    'Create new folder to save the new files in
    DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
    FolderName = Sourcewb.Path & "\" & Sourcewb.Name & " " & DateString
    MkDir FolderName

    'Copy every visible sheet to a new workbook
    For Each sh In Sourcewb.Worksheets

        'If the sheet is visible then copy it to a new workbook
        If sh.Visible = -1 Then
            sh.Copy

            'Set Destwb to the new workbook
            Set Destwb = ActiveWorkbook

            'Determine the Excel version and file extension/format
            With Destwb
                If Val(Application.Version) < 12 Then
                    'You use Excel 97-2003
                    FileExtStr = ".xls": FileFormatNum = -4143
                Else
                    'You use Excel 2007-2013
                    If Sourcewb.Name = .Name Then
                        MsgBox "Your answer is NO in the security dialog"
                        GoTo GoToNextSheet
                    Else
                        Select Case Sourcewb.FileFormat
                        Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
                        Case 52:
                            If .HasVBProject Then
                                FileExtStr = ".xlsm": FileFormatNum = 52
                            Else
                                FileExtStr = ".xlsx": FileFormatNum = 51
                            End If
                        Case 56: FileExtStr = ".xls": FileFormatNum = 56
                        Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
                        End Select
                    End If
                End If
            End With

            'Save the new workbook and close it
            With Destwb
                .SaveAs FolderName _
                      & "\" & Destwb.Sheets(1).Name & FileExtStr, _
                        FileFormat:=FileFormatNum
                .Close False
            End With

        End If
GoToNextSheet:
    Next sh

    MsgBox "You can find the files in " & FolderName

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

```

Any help really gratefully received as I have been at this now for three days...

----

`sh.Copy` will create a new `Workbook` out of the `ActiveSheet` by default. [Take a look at the docs to see that](https://msdn.microsoft.com/en-us/library/office/ff837784.aspx). This is the same as right clicking on the worksheet tab and selecting `Move or Copy...` and then copying to a new workbook.

If you want to only keep values and source formatting, you can do this by manually creating a new `Workbook`, copying the cells, and using `PasteSpecial` to only get the values and formats in the new workbook.

This is a quick drop in change for your code. Replace the lines:

```
sh.Copy

'Set Destwb to the new workbook
Set Destwb = ActiveWorkbook

```

With something like this:

```
Dim Destwb As Workbook
Set Destwb = Workbooks.Add

Dim sh_copy As Worksheet
Set sh_copy = Destwb.Worksheets(1)

sh.Cells.Copy

sh_copy.Cells.PasteSpecial xlPasteValuesAndNumberFormats
sh_copy.Cells.PasteSpecial xlPasteFormats

sh_copy.Name = sh.Name

Application.CutCopyMode = False

```

The idea is that you:

*   manually create the new `Workbook`
*   grab a reference to the first `Worksheet` in the new book. Note that this workbook may have multiple sheets depending on your default settings. If so, either change the default number of sheets, or add some code to delete the extra ones.
*   copy the cells from the original sheet and use `PasteSpecial` to get the values and numbers formats and then another call to get the column width info
*   set the name to the original name since you will get a default name at first
*   remove the copy/paste info
