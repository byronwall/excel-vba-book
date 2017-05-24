# SO item 048
I want to create a macro that copies the top row of formulas and continue dropping it down the worksheet until one of the formulas in the previous row returns a blank.

Here is my code:

```
Range("C8:V8").Select
Selection.Copy
Do
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Past
Loop Until ActiveCell.offest(-1, 15) = ""

```

Any thoughts on why I keep getting an error?

----

If you want to check all of the cells for a possible blank, you can use the VBA version of `COUNTBLANK` for that.

```
Loop Until Application.CountBlank(ActiveCell.Offset(-1).Resize(,15)) > 0

```

The call to `Resize` is needed to get a range that includes all of the cells for all 15 columns.
