# SO item 006
I have a column of cells (Range A2:A10) which contain names of students. For each student, I have a chart titled after their name which track their performances in another sheet. I would like to change the chart background to light red colour if their names appear in the column of cells.

----

This is a fairly straight forward combination of ideas. You need to iterate through the charts, check the title against a list, and change the background color. The code below is an example to show the idea.

When iterating through Charts on a sheet, you start with the ChartObjects method. The ChartObject contains a reference to the actual Chart where you can get the Title and change the background. Note that checking the Title of a Chart without one will throw an error, so I start with a check to Chart.HasTitle.

I am using Application.Match to check if the range contains the title. This will return an error if it is not found, so I am checking for that error.

Finally, if the match exists, you change the background of the Chart through a lengthy list of properties. If you want to change a different part of the chart, record a macro to find the right property.

```
Sub ColorBasedOnTitle()

    Dim chtObj As ChartObject
    Dim sht As Worksheet
    Dim rng_students As Range

    'assume active sheet, change if not
    Set sht = ActiveSheet

    'need to set a reference to the list of names... named range is probably prefered here
    Set rng_students = sht.Range("B3:B6")

    'loop through all charts on sheet
    For Each chtObj In sht.ChartObjects

        'if chart has title, check its value
        Dim title As String
        If chtObj.Chart.HasTitle Then
            title = chtObj.Chart.ChartTitle.Text

            'use Match to see if title is in list of names
            Dim search As Variant
            search = Application.Match(title, rng_students, 0)

            'see if student is in list, change background if so
            If Not IsError(search) Then
                chtObj.Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
            End If
        End If
    Next chtObj
End Sub

```

Here is a picture of my Excel instance so you can see the result. Note that my Chart titled "F" is outside the range of checked names which are highlighted grey for emphasis.

![charts after code runs](https://i.stack.imgur.com/Y7kzZ.png)
