# SO item 050
Say we have some long formula saved in cell **A1**:

```
=SomeArrayFunction(
IF(SUM(D3:D6)>1,"A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X 01",
"part_one"),
IF(SUM(D3:D6)>1,"A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X 02",
IF(SUM(D3:D6)>1,"A-B-C-D-E-F-G-H-I-J-K-L-M-N-O-P-Q-R-S-T-U-V-W-X 03",
"part_two"))
)

```

which uses the following VBA function

```
Public Function SomeArrayFunction(sOne As String, sTwo As String) As Variant
    Dim V() As Variant
    ReDim V(1 To 2, 1 To 1)
    V(1, 1) = sOne
    V(2, 1) = sTwo
    SomeArrayFunction = V
End Function

```

returning a 2×1 array.

* * *

Now when I call this VBA function

```
Public Sub EvaluateFormula()
    Dim vOutput As Variant

    vOutput = Application.Evaluate(Selection.Formula)

    If VarType(vOutput) >= vbArray Then
        MsgBox "Array:" & vbCrLf & vOutput(1, 1) & vbCrLf & vOutput(2, 1)
    Else
        MsgBox "Single Value: " & vbCrLf & vOutput
    End If
End Sub

```

while having selected cell **A1** I get an error, because Application.Evaluate cannot handle formulas with more than 255 characters (e.g. see [VBA - Error when using Application.Evaluate on Long Formula](http://stackoverflow.com/questions/30267826/vba-error-when-using-application-evaluate-on-long-formula)). On the other hand, if I write

```
vOutput = Application.Evaluate(Selection.Address)

```

instead (as proposed in the link above), then it works just fine. Except for the fact that the array is not being recgonised anymore, i.e. _MsgBox "Single Value: "_ is called instead of _MsgBox "Array:"_.

So my question is: How can I evaluate long formulas (which return arrays) using VBA?

* * *

**Edit:** Let me stress that I need this to work when I only select the **one cell** that conains the formula (not a region or several cells). And I have not entered it as an array formula (i.e. no curly brackets): ![enter image description here](https://i.stack.imgur.com/FDvrh.png)

* * *

**Edit2:** Let me answer the why: my current work requires me to have a long list of such large formulas in a spreadsheet. And since they are organised in a list every such formula can only take up one cell. In almost all cases the formulas return single values (and hence one cell is sufficient to store/display the output). However, when there is an internal error in evaluating the formula, the formula returns an error message. These error messages are usually quite long and are therefore returned as an array of varying size (depending on how long the error message is). So my goal was to write a VBA function that would first obtain and then output the full error message for a given selected entry from the list.

----

I believe that `Application.Evaluate` will return a result that matches the size of the input address. I suspect that your `Selection` is a single cell so it is returning a single value.

If instead you call it with `Selection.CurrentArray.Address` you will get an answer that is the same size as the correct array.

**Picture of VBA and Excel**

![enter image description here](https://i.stack.imgur.com/1h2Rj.png)

**Code to test with**

```
Public Function Test() As Variant

    Test = Array(1, 2)

End Function

Sub t()

    Dim a As Variant

    a = Application.Evaluate(Selection.CurrentArray.Address)

End Sub

```

**Edit**, based on comments here is a way evaluate this off sheet by creating a new sheet. I am using a cut/paste approach to ensure the formulas all work the same. This probably works better if cells don't reference the cut one. It will technically not break any other cells though since I am using cut/paste.

In the code below, I had an array formula in cell `J2` it referenced several other cells. It is expanded to have 3 rows and then the `Evaluate` call is made. That returns an array like you want. It then shrinks it down to one cell and moves it back.

I have tested this for a simple example. I have no idea if it works for the application you have in mind.

```
Sub EvaluateArrayFormulaOnNewSheet()

    'cut cell with formula
    Dim str_address As String
    Dim rng_start As Range
    Set rng_start = Sheet1.Range("J2")
    str_address = rng_start.Address

    rng_start.Cut

    'create new sheet
    Dim sht As Worksheet
    Set sht = Worksheets.Add

    'paste cell onto sheet
    Dim rng_arr As Range
    Set rng_arr = sht.Range("A1")
    sht.Paste rng_arr

    'expand array formula size.. resize to whatever size is needed
    rng_arr.Resize(3).FormulaArray = rng_arr.FormulaArray

    'get your result
    Dim v_arr As Variant
    v_arr = Application.Evaluate(rng_arr.CurrentArray.Address)

    ''''do something with your result here... it is an array

    'shrink the formula back to one cell
    Dim str_formula As String
    str_formula = rng_arr.FormulaArray

    rng_arr.CurrentArray.ClearContents
    rng_arr.FormulaArray = str_formula

    'cut and paste back to original spot
    rng_arr.Cut

    Sheet1.Paste Sheet1.Range(str_address)

    Application.DisplayAlerts = False
    sht.Delete
    Application.DisplayAlerts = True

End Sub

```
