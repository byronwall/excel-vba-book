# SO item 079
I have a set of data where each item has a 2D array of information corresponding to it. I'd like to create a 3D array where the first dimension is the item name and the second and third dimensions are taken up by the info corresponding to the item.

I can read the data for each item into a 2D array, but I can't figure out how to get the 2D array read into the 3D array.

I know the sizes of all the dimensions so I can create an array of fixed size before I begin the reading and writing process.

I'd like to do this by looping only through the names of the items and not looping through every cell of every 2D array.

It is easy to get the 2D arrays read in to an ArrayList but I want to be able to name the items and be able to read these back in to excel and it seems difficult to do with an ArrayList.

The question is: how do I read a 2D selection from excel into a 3D fixed sized array in VBA?

----

Here is an example of each approach: array of arrays or `Dictionary` of arrays. The **Dictionary approach is considerably easier** than the array of arrays if what you want is keyed lookup of values. There might be merits to the array of arrays in other cases.

This is dummy code with no real purpose but to show a couple things: grabbing a single value and an array of values. I am building a 2D array of values by grabbing some `Range` data which is easy and 2D. From there I build up the arrays of these values and then put them into the relevant data structure. Then I poke at the data structure to get some values out of it.

**Array of Arrays approach** is shown first (and outputs on the left of the picture).

```
Sub ArraysOfArrays()

    Dim arrA() As Variant
    Dim arrB() As Variant

    'wire up a 2-D array
    arrA = Range("B2:D4").Value
    arrB = Range("F2:H4").Value

    Dim arrCombo() As Variant
    ReDim arrCombo(2, 1) As Variant

    'name and give data
    arrCombo(0, 0) = "arrA"
    arrCombo(1, 0) = arrA

    'add more elements
    ReDim Preserve arrCombo(2, 2)

    arrCombo(0, 1) = "arrB"
    arrCombo(1, 1) = arrB

    'output a single result
    'cell(2,2) of arrA
    Range("B6") = arrCombo(1, 0)(2, 2)

    Dim str_search As String
    str_search = "arrB"

    'iterate through and output arrB to cells
    Dim i As Integer
    For i = LBound(arrCombo, 1) To UBound(arrCombo, 1)
        If arrCombo(0, i) = str_search Then
            Range("B8").Resize(3, 3).Value = arrCombo(1, i)
        End If
    Next i
End Sub

```

Couple key points here:

*   You can only expand the array using `ReDim`. `ReDim` is very particular that you only [change the last dimension of the array](http://stackoverflow.com/a/13184027/4288101) when used with `Preserve`. Since I need one of them to track the number of entries, I do that in the second index which is... unnatural. If you know the size in advance, this painful step is skipped.
*   My final array is a 2xN array where the 2 contains a name and a YxZ array of data.
*   In order to find a given array in the mix, you have to iterate through them all.

**Dictionary of Arrays** is far less code and more elegant. Be sure to add the reference `Tools->References` in the VBA editor to Microsoft Scripting Runtime.

```
Sub DictionaryOfArrays()

    Dim dict As New Scripting.Dictionary

    'wire up a 2-D array
    arrA = Range("B2:D4").Value
    arrB = Range("F2:H4").Value

    dict.Add "arrA", arrA
    dict.Add "arrB", arrB

    'get a single value
    Range("F6") = dict("arrB")(2, 2)

    'get a array of values
    Range("F8").Resize(3, 3) = dict("arrA")

End Sub

```

**Picture of input data and results**

![data and results](https://i.stack.imgur.com/ApX58.png)

Data to copy if you want it (paste in `B1`)

```
a               b       
1   2   3       10  11  12
4   5   6       13  14  15
7   8   9       16  17  18

```
