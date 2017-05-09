```vb
Public Sub UpdateFromChartSeries(targetSeries As series)
    '---------------------------------------------------------------------------------------
    ' Procedure : UpdateFromChartSeries
    ' Author    : @byronwall
    ' Date      : 2015 11 09
    ' Purpose   : Reads the series info from a Series and stores it in the class
    '---------------------------------------------------------------------------------------
    '
    
    'this will work for the simple case where all items are references
    Const FIND_STRING As String = "SERIES("
    Const COMMA As String = ","
    Const CLOSE_BRACKET As String = ")"
    
    Set series = targetSeries

    Dim targetForm As Variant

    '=SERIES("Y",Sheet1!$C$8:$C$13,Sheet1!$D$8:$D$13,1)

    'pull in teh formula
    targetForm = targetSeries.Formula

    'uppercase to remove match errors
    targetForm = UCase(targetForm)

    'remove the front of the formula
    targetForm = Replace(targetForm, FIND_STRING, vbNullString)
    
    'find the first foundPosition
    Dim foundPosition As Long
    foundPosition = InStr(targetForm, COMMA)

    If foundPosition > 1 Then
        'need to catch an error here if a text name is used instead of a valid range
        On Error Resume Next
        Set Me.name = Range(left(targetForm, foundPosition - 1))
        If Err <> 0 Then pName = left(targetForm, foundPosition - 1)
        On Error GoTo 0
    End If

    'pull out the title from that
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, COMMA)

    If foundPosition > 1 Then Set Me.XValues = Range(left(targetForm, foundPosition - 1))
 
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, COMMA)
    Set Me.Values = Range(left(targetForm, foundPosition - 1))
    targetForm = Mid(targetForm, foundPosition + 1)

    foundPosition = InStr(targetForm, CLOSE_BRACKET)
    Me.SeriesNumber = left(targetForm, foundPosition - 1)

    Me.ChartType = targetSeries.ChartType
End Sub
```