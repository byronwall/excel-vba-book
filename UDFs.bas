Attribute VB_Name = "UDFs"
Option Explicit


Public Function RandLetters(ByVal letterCount As Long) As String
    '---------------------------------------------------------------------------------------
    ' Procedure : RandLetters
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : UDF that generates a sequence of random letters
    '---------------------------------------------------------------------------------------
    '
    Dim letterIndex As Long
    
    Dim letters() As String
    ReDim letters(1 To letterCount)
    
    For letterIndex = 1 To letterCount
        letters(letterIndex) = chr(Int(Rnd() * 26 + 65))
    Next
    
    RandLetters = Join(letters(), "")
    
End Function

Public Function ConcatRange(rngCells As Range, strDelim As String) As String
    Dim cellCount As Long
    
    cellCount = rngCells.CountLarge
    
    Dim arrValues As Variant
    ReDim arrValues(1 To cellCount)
    
    Dim index As Long
    index = 1
    
    Dim rngCell As Range
    For Each rngCell In rngCells
        arrValues(index) = rngCell
        
        index = index + 1
    Next
    
    ConcatRange = Join(arrValues, strDelim)
End Function


Public Function ConcatArr(rngCells As Variant, strDelim As String) As String
    Dim cellCount As Long
    
    cellCount = UBound(rngCells, 1)
    
    Dim arrValues As Variant
    ReDim arrValues(1 To cellCount)
    
    Dim index As Long
    index = 1
    
    Dim rngCell As Variant
    For Each rngCell In rngCells
        arrValues(index) = rngCell
        
        index = index + 1
    Next
    
    ConcatArr = Join(arrValues, strDelim)
End Function
