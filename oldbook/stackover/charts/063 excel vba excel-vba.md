# SO item 063
I have some code which changes all texts in a chart object to a specific size and font. Thing is the first time one runs the code it works like a charm. But if I change any part of the text within the chart and then re-run the code nothing happens.

E.g I run the code, then change the title heading to size 15 and font arial, then rerun macro, and nothing happens.

What can be wrong?

My code

```
With Selection

        .Format.TextFrame2.TextRange.Font.Size = 10
        .Format.TextFrame2.TextRange.Font.Name = "Times New Roman"
        .Format.TextFrame2.TextRange.Font.Bold = msoFalse

        End With

```

----

When you apply the fonts/sizes to the `ChartArea` in order to propagate them down to the individual pieces, Excel stores that info at the `CharArea` level. If you make a change to the one of the components (`ChartTitle`, `Axis`, etc.) and try to run your code again, there is no change on the `ChartArea`. Seems that Excel does not propagate those changes "back up". This makes sense since now the different items are styled differently.

The easiest way to deal with this is to reset the styles before you make your changes. `ClearToMatchStyle` applied to the `Chart` (i.e. `ActiveChart` or `Selection.Parent` in your context) will do it. It appears it will also make the change if you use a different font size or actually make a change to one of the `ChartArea.Format` properties (e.g. `Size`, `Name`, etc.).

**Code for the reset option**

```
ActiveChart.ClearToMatchStyle

With Selection

    .Format.TextFrame2.TextRange.Font.Size = 12
    .Format.TextFrame2.TextRange.Font.Name = "Times New Roman"
    .Format.TextFrame2.TextRange.Font.Bold = msoFalse

End With 

```
