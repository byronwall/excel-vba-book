```vb
Sub MakeSeveralBoxesWithNumbers()

    Dim shp As Shape
    Dim sht As Worksheet

    Dim rng_loc As Range
    Set rng_loc = Application.InputBox("select range", Type:=8)

    Set sht = ActiveSheet

    Dim int_counter As Long

    For int_counter = 1 To InputBox("How many?")

        Set shp = sht.Shapes.AddTextbox(msoShapeRectangle, rng_loc.left, _
                                        rng_loc.top + 20 * int_counter, 20, 20)

        shp.Title = int_counter

        shp.Fill.Visible = msoFalse
        shp.Line.Visible = msoFalse

        shp.TextFrame2.TextRange.Characters.Text = int_counter

        With shp.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
            .Solid
        End With

    Next

End Sub
```