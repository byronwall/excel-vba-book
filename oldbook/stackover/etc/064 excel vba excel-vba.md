# SO item 064
I'm a beginner at VBA and I've been scratching my head all day trying to figure out what's wrong with my code:

```
Sub DataTransfer():
    Dim position As Integer, location (1 To 9) As String
    location(1) = "BC"
    location(2) = "Calgary"
    location(3) = "Edmonton"
    location(4) = "Major Projects"
    location(5) = "Minneapolis"
    location(6) = "Saskatchewan"
    location(7) = "Seattle"
    location(8) = "Toronto"
    location(9) = "Winnipeg"

    For position = 1 To 9
        Worksheets(location(position)).Select
        Cells(1, 2).Value = location(position)
    Next position
End Sub

```

Edit: Sorry about the ambiguity of my question. What I ultimately want to do is actually to be able to change that third last line (that writes the city name to the worksheets) to any function I want so as to modify the worksheets as I see fit. This is actually part of a larger subroutine that I broke out to troubleshoot the problem. These worksheets are interspersed between other worksheets so unfortunately, @nutsch's solution won't really achieve what I want (but thanks either way).

The biggest problem I have with this is that this exact code would sometimes work as intended and other times return the "subscript out of range" error on the fourth last line.

----

I assume you are trying to put the sheet name into the worksheet it corresponds to? If so, the problem is that you are using `Select` instead of `Activate` to give the `Worksheet` focus.

The becomes a problem because you are using `Cells` without a qualifier so it refers to the `ActiveSheet` which has not been changed in your code.

Two solutions:

*   Use `Activate` which will make `Cells` work like you want.
*   Qualify the call to `Cells` by preceding it with a sheet object.

**Option 1**

```
Sub DataTransfer():
    Dim position As Integer, location(1 To 9) As String
    location(1) = "BC"
    location(2) = "Calgary"
    location(3) = "Edmonton"
    location(4) = "Major Projects"
    location(5) = "Minneapolis"
    location(6) = "Saskatchewan"
    location(7) = "Seattle"
    location(8) = "Toronto"
    location(9) = "Winnipeg"

    For position = 1 To 9
        Worksheets(location(position)).Activate
        Cells(1, 2).Value = location(position)
    Next position
End Sub

```

**Option 2**

```
Sub DataTransfer():
    Dim position As Integer, location(1 To 9) As String
    location(1) = "BC"
    location(2) = "Calgary"
    location(3) = "Edmonton"
    location(4) = "Major Projects"
    location(5) = "Minneapolis"
    location(6) = "Saskatchewan"
    location(7) = "Seattle"
    location(8) = "Toronto"
    location(9) = "Winnipeg"

    For position = 1 To 9
        'this line really does nothign now
        Worksheets(location(position)).Select

        Worksheets(location(position)).Cells(1, 2).Value = location(position)
    Next position
End Sub

```

I prefer Option 2 because it does not require activating the `Worksheet` first. It should always be your goal to avoid Activating and Selecting when possible. They slow things down and make code the opposite of robust.

Finally, as pointed out by @nutsch there are easier ways to do something simialr to this, but it's worth knowing why your code doesn't work.
