# SO item 116
I receive excel workbooks from various customers containing lists of employees with gender and date of birth columns. They arrive in varying formats. I am creating an excel worksheet that I will add to each of these workbooks which will allow me to apply a series of formulas to the customer data. The problem is that I do not know the location of the date of birth data in advance.

I want a macro to prompt me to select the range for the date of birth values and then place that data range into a cell (F31) on my worksheet so that I can pull it into my other formulas.

The code I created below works, but it does not pull the worksheet tab name along with the range. How can I get the worksheet name along with the cell range?

```
Sub ChooseDOBRange()
    Dim rng As Range

    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    rng.Copy

    Worksheets("COVER SHEET").Range("F31") = rng.Address    
End Sub 

```

----

There are two ways to get this depending on whether or not you want the `Workbook` name to be in the formula also. I am testing these in the Immediate Window.

**Method 1** uses `.Address` with the `External` parameter set to `True`

```
?ActiveCell.Address(,,,True)

```

> [Book2]Sheet1!$A$1

**Method 2** uses the `.Address` along with the sheet name from `Range.Parent.Name` where `Parent` refers to the `Worksheet` for a `Range`

```
?"'" & ActiveCell.Parent.Name & "'!" & ActiveCell.Address

```

> 'Sheet1'!$A$1
