# SO item 051
I have a project where I need to count the number of values depending on two factors. For every cell in N, which = EE Only, the corresponding cell in column O = 1\. The issue is, for cells where the values are EE+FAM or EE+SP, they need to count the number of instances where those values occur between the corresponding values of "M" in column L, as shown in the picture. The picture shows what the output should be.

![enter image description here](https://i.stack.imgur.com/PNyJf.png)

The code I have thus far can put a 1 if the value in column N is "EE Only", and 2 if the value is something else. I am not sure how to build in a second set of conditions that checks for the values in column L . I am pretty new to VBA, so any help is appreciated. Here is the code.

```
Sub CountDependents()
    Dim DepCount As Long
    Dim LastRowConsole As Long
    Dim m As Integer
    Dim l As Integer
    Dim j As Integer
    Dim k As Integer
    Dim ws As Worksheet

    Set ws = Sheets("Audit page")

    m = 1
    For l = 4 To 15
        If Cells(l, "N") <> "EE Only" Then
            m = m + 1
        End If

        Cells(l, "O").Value = m
        m = 1
    Next l

    k = 1
    For j = 4 To 15
        If Cells(j, "L") = "M" Then
            k = k + 1
        End If

        Cells(j, "O").Value = Cells(j + Cells(j, "O") / 2, "O").Value
        j = 1
    Next j               
End Sub

```

----

If you want to process the names column, here is some simple code which will take care of it. The idea is that you "build" a `Range` using `Union` and then count those cells when you hit a "switch". You can then output that count wherever you want.

In my dummy example below, I am using `d` instead of `dependent`. I am also outputting right next to the names column. You can modify these to suit your application.

```
Sub CountPeople()

    Dim rng_cell As Range
    Dim rng_top As Range

    Set rng_top = Range("K3")

    Dim bool_first As Boolean
    bool_first = True

    For Each rng_cell In Range(rng_top, rng_top.End(xlDown))

        If rng_cell <> "d" And Not bool_first Then
            'starting new person, output result from last to top of last person
            rng_top.Offset(, 1) = rng_top.Cells.Count

            'new top cell is current cell
            Set rng_top = rng_cell

        Else
            'keep growing range if we see a dependent
            Set rng_top = Union(rng_top, rng_cell)
        End If

        bool_first = False
    Next

    'handle the last person
    rng_top.Offset(, 1) = rng_top.Cells.Count

End Sub

```

**Results**, note that the output went into the yellow cells.

![results of VBA](https://i.stack.imgur.com/pHBt3.png)

**Rough description of code**

*   Picking a starting point.
*   Iterate through all names including that cell and below it.
*   If the current cell is equal to `d` then continue building a range
*   Once we see something other than `d`, we need to count the previous cells and output the count in the column next door.
*   There is a line to handle the first case (it should not equal `d` but we want to count it and keep going) and also the last name which will finish outside the loop.
