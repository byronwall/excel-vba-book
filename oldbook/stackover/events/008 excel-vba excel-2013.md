# SO item 008
I wonder whether someone could help me please.

I'm using the code below to dynamically create a list of files from a given folder.

In column E for each row of the list there is a link 'Click Here to Open' which allows the user to open each file.

But I'm now looking to change this so rather than opening the file, the link will open a 'Save' dialog which allows the user the file to a user selected folder, and I must admit this issue has had me baffled for over a week now.

```
Public Sub ListFilesInFolder(SourceFolder As Scripting.folder, IncludeSubfolders As Boolean)

    Dim LastRow As Long

    On Error Resume Next
    For Each FileItem In SourceFolder.Files
        ' display file properties
        Cells(iRow, 3).Formula = iRow - 12
        Cells(iRow, 4).Formula = FileItem.Name
        Cells(iRow, 5).Select
        Selection.Hyperlinks.Add Anchor:=Selection, Address:= _
        FileItem.Path, TextToDisplay:="Click Here to Open"
        iRow = iRow + 1 ' next row number

        With ActiveSheet
            LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
            LastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        End With

        For Each Cell In Range("C13:E" & LastRow) ''change range accordingly
            If Cell.Row Mod 2 = 1 Then ''highlights row 2,4,6 etc|= 0 highlights 1,3,5
                Cell.Interior.Color = RGB(232, 232, 232) ''color to preference
            Else
                Cell.Interior.Color = RGB(141, 180, 226) 'color to preference or remove
            End If
        Next Cell
    Next FileItem

    If IncludeSubfolders Then
        For Each SubFolder In SourceFolder.SubFolders
            ListFilesInFolder SubFolder, True
        Next SubFolder
    End If
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
End Sub

```

I've tried using the command 'Application.Dialogs(xlDialogSaveAs).Show' in every row of the code but I cannot get this to work because all it does is ask the user to save the file as it creates the list.

I just wondered whether someone could possibly look at this please and let me know where I've gone wrong.

Many thanks and kind regards

Chris

----

Here is the relevant code for copying a file from a given spot to a destination folder selected by the user. I wrapped it in a Worksheet event for `FollowHyperlink` since it sounds like you are doing this based on a click.

```
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Dim FSO
    Dim sFile As String
    Dim sDFolder As String

    'path to file to copy, you will want to point this at a cell range
    'this assume a single cell is selected
    sFile = Target.Range.Value

    'destination folder
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

    fldr.AllowMultiSelect = False
    fldr.Show

    'add the end slash for the copy operation
    sDFolder = fldr.SelectedItems(1) & "\"

    'FSO object to copy the file... True below overwrites if needed
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile (sFile), sDFolder, True

End Sub

```
