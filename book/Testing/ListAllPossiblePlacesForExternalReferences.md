```vb
Public Sub ListAllPossiblePlacesForExternalReferences()

    'search through chart formulas
    Debug.Print "Checking chart series formulas..."
    Dim chtObj As ChartObject
    For Each chtObj In Chart_GetObjectsFromObject(ActiveSheet)
        Dim ser As series
        For Each ser In chtObj.Chart.SeriesCollection
            
            Dim strForm As String
            strForm = ser.Formula
            
            If InStr(strForm, "[") Then
                Debug.Print strForm
            End If
        Next
    Next
    
    'search in data validation
    Dim sht As Worksheet
    Dim rng As Range
    Debug.Print "Checking data validation formulas..."
    For Each sht In Worksheets
        For Each rng In sht.UsedRange
            Dim strVal As String
            strVal = "!"
            On Error Resume Next
            strVal = rng.Validation.Formula1
            On Error GoTo 0
            
            If strVal <> "!" Then
                If InStr(strVal, "[") Then
                    Debug.Print rng.Address(False, False, , True) & strVal
                    'rng.Activate
                End If
            End If
        Next
    Next
    
    'search in conditional formatting
    Debug.Print "Checking conditional formatting formulas..."
    For Each sht In Worksheets
        For Each rng In sht.UsedRange
            Dim condFormat As FormatCondition
            For Each condFormat In rng.FormatConditions
                'get the formulas
        
                strVal = "!"
                On Error Resume Next
                strVal = condFormat.Formula1
                On Error GoTo 0
            
                If strVal <> "!" Then
                    If InStr(strVal, "[") Then
                        Debug.Print rng.Address(False, False, , True) & strVal
                        'rng.Activate
                    End If
                End If
            Next
        Next
    Next
End Sub
```