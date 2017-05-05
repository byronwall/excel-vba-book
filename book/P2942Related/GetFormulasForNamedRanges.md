```vb
Public Sub GetFormulasForNamedRanges()

    Dim n As Integer
    n = FreeFile()
    Open "C:\Documents\TDA\2942\Mass balance\formulas.ini" For Output As #n

    Dim namedRange As name
    For Each namedRange In ActiveWorkbook.Names
        
        Dim hasForm As Boolean
        hasForm = False
        
        On Error Resume Next
        hasForm = namedRange.RefersToRange.HasFormula
        On Error GoTo 0
        
        If namedRange.Visible And namedRange.name <> "SELF" And InStr(namedRange.name, "IGNORE") = 0 Then
            
            If hasForm Then
                'this allows for self reference
                Dim strFormula As String
                strFormula = namedRange.RefersToRange.Formula
            
                Range("J90").Formula = strFormula
                On Error Resume Next
                strFormula = Range("K90").Value
                
                If Err.Number <> 0 Then
                    Debug.Print "Error"
                End If
                
                On Error GoTo 0
            
                strFormula = Replace(strFormula, "SELF", namedRange.name)
                
                If namedRange.name = "FT_601_slpm" Then
                    strFormula = strFormula & "*SIGN(FT_601_dp)"
                End If
                
                Dim deadband As Double
                deadband = 1000000
                
                If InStr(namedRange.name, "capture") Then
                    deadband = 0.1
                End If
                
                
                strFormula = Replace(strFormula, "*", " * ")
                strFormula = Replace(strFormula, "-", " - ")
                strFormula = Replace(strFormula, "+", " + ")
                strFormula = Replace(strFormula, "/", " / ")
                strFormula = Replace(strFormula, "ABS(", "abs(")
                strFormula = Replace(strFormula, "_", "-")
                
                Print #n, ";" & namedRange.name & " = " & namedRange.Comment
                
                Print #n, "[" & Replace(namedRange.name, "_", "-") & "]"
                
                Print #n, "type = float"
                Print #n, "math = " & strFormula
                Print #n, "daq_ch ="
                Print #n, "daq_io ="
                Print #n, "daq_cal ="
                Print #n, "filter ="
                Print #n, "mb_ip ="
                Print #n, "mb_ch ="
                Print #n, "mb_type ="
                Print #n, "mb_order ="
                Print #n, "mb_io ="
                Print #n, "log_series ="
                Print #n, "daq_type ="
                Print #n, "log_type = float"
                Print #n, "deadband = " & deadband
            
                Print #n, ""
            End If
        End If
    Next
    
    Close #n
    

End Sub
```