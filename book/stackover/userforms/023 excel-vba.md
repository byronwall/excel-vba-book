# SO item 023
Is there a way to display an item in a combobox list but give it a different value?

Say I have "1-03" in my list representing 1'3" but I want the value of "1.25" assigned to it so my formulas will calculate correctly.

```
Private Sub UserForm_Activate()

'Values for cmbxSpans
cmbxSpans.AddItem "1-03"

End Sub

```

----

No. Excel does not provide additional properties on the items (entries in `List(i)`) to differentiate display value and "actual" value. The items in a `ListBox` are stored as `Strings`. If you want a two-fold representation (i.e `1-03 = 1.25`), you will have to handle conversion when items are added/read.

Here is one such set of conversions based on your `1-03` example.

```
Private Sub CommandButton1_Click()

    ListBox1.AddItem "1-03"

    'read item
    Dim height As Double
    height = FormatToNumber(ListBox1.List(0))

    'do some math
    height = height * 2

    'add it back it
    ListBox1.AddItem NumberToFormat(height)

End Sub

Function FormatToNumber(str_feet_inch As String) As Double
    Dim values As Variant

    'split based on -
    values = Split(str_feet_inch, "-")

    'do the math, using 12# to ensure double result
    FormatToNumber = values(0) + values(1) / 12#
End Function

Function NumberToFormat(val As Double) As String

    Dim str_feet As String
    Dim str_inch As String

    str_feet = Format(Int(val), "0")
    str_inch = Format((val - Int(val)) * 12, "00")

    NumberToFormat = str_feet & "-" & str_inch

End Function

```
