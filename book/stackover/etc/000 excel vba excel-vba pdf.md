# SO item 000
I'm learning VBA as a means to speed up some processes at my work. I have to send rooming lists to properties in PDF format generated form excel. I have the following code that works great, but I get an error message whenever one of the worksheets is hidden. I have to hide sheets often as properties change from trip to trip.

I want to PDF the Worksheets from the 4th sheet to the sheetname "Post". Whenever I hide a sheet in between these I get the following error message "Run-time error '5': Invalid procedure call or argument"

Here's the code:

```
Sub SaveAllPDF()
Dim I As Integer
Dim Fname As String
Dim TabCount As Long

TabCount = Sheets("Post").Index

' Begin the loop.

For I = 4 To TabCount
Sheets(I).Activate
With ActiveSheet
Fname = .Range("C15") & " " & .Range(" B1")
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
"C:\Users\brandon.ford\Desktop\Operation Automated\" & Fname,     
Quality:=xlQualityStandard, _
IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End With
Next I
End Sub

```

Anyone have any ideas how to remedy this problem so the 'For I = 4 to TabCount' ignores any hidden tabs? Any help would be MUCH appreciated, I've been trying to tackle this problem for a long time and do not have much VBA knowledge.

----

With your current loop, this is best avoided by simply checking that the WorkSheet is visible before trying to export. The Visible property contains this info. The value should be xlSheetVisible if the sheet is visible.

Here is the full code with the check:

```
Sub SaveAllPDF()
Dim I As Integer
Dim Fname As String
Dim TabCount As Long

TabCount = Sheets("Post").Index

' Begin the loop.

For I = 4 To TabCount
Sheets(I).Activate
With ActiveSheet
    If .Visible = xlSheetVisible Then

        Fname = .Range("C15") & " " & .Range(" B1")
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        "C:\Users\brandon.ford\Desktop\Operation Automated\" & Fname,
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    End If
End With
Next I
End Sub

```
