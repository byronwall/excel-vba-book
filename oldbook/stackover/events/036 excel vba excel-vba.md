# SO item 036
I am looking around for a formula to hide certain rows based on certain cell inputs. In cell `C5` I have a drop-down selection of "`Corporates`" and "`Projects`". In cell `C8` I have a drop-down selection of "`High`", "`Medium`", and "`Low`". In cell `H6` I have the formula `=C5&C8`. The macro I have is as follows:

```
Private Sub Worksheet_Change(ByVal Target As Range)
If Target.Address(False, False) = "H6" Then
    Select Case Target.Value
        Case "CorporatesHigh": Rows("21:33").Hidden = True: Rows("12:20").Hidden = False
        Case "CorporatesMedium": Rows("21:33").Hidden = True: Rows("12:20").Hidden = False
        Case "CorporatesLow": Rows("25:33").Hidden = True: Rows("12:24").Hidden = False
        Case "ProjectsHigh": Rows("25:28").Hidden = False: Rows("29:33").Hidden = True: Rows("12:24").Hidden = True
        Case "ProjectsMedium": Rows("25:28").Hidden = False: Rows("29:33").Hidden = True: Rows("12:24").Hidden = True
        Case "ProjectsLow": Rows("25:33").Hidden = False: Rows("12:24").Hidden = True
        Case "": Rows("12:33").Hidden = False
        Case "Corporates": Rows("12:33").Hidden = False
        Case "Projects": Rows("12:33").Hidden = False
        Case "High": Rows("12:33").Hidden = False
        Case "Medium": Rows("12:33").Hidden = False
        Case "Low": Rows("12:33").Hidden = False
    End Select
End If
End Sub

```

The macro works when I click into `H6` but I need it to work when cells `C5` or `C8` is changed.

----

The `If-Not-Intersect-Is Nothing` pattern is the preferred way to check if the changed cells(s) are inside a given range of cell(s).

To handle all three of your cells, you can use the following. You can easily add or remove cells to the `Range` if you want.

```
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("H6,C5,C8")) Is Nothing Then
        Select Case Range("H6").Value
            Case "CorporatesHigh": Rows("21:33").Hidden = True: Rows("12:20").Hidden = False
            Case "CorporatesMedium": Rows("21:33").Hidden = True: Rows("12:20").Hidden = False
            Case "CorporatesLow": Rows("25:33").Hidden = True: Rows("12:24").Hidden = False
            Case "ProjectsHigh": Rows("25:28").Hidden = False: Rows("29:33").Hidden = True: Rows("12:24").Hidden = True
            Case "ProjectsMedium": Rows("25:28").Hidden = False: Rows("29:33").Hidden = True: Rows("12:24").Hidden = True
            Case "ProjectsLow": Rows("25:33").Hidden = False: Rows("12:24").Hidden = True
            Case "": Rows("12:33").Hidden = False
            Case "Corporates": Rows("12:33").Hidden = False
            Case "Projects": Rows("12:33").Hidden = False
            Case "High": Rows("12:33").Hidden = False
            Case "Medium": Rows("12:33").Hidden = False
            Case "Low": Rows("12:33").Hidden = False
        End Select
    End If
End Sub

```
