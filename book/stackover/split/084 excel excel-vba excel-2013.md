# SO item 084
I'm Using Microsoft Excel 2013.

I have a lot of data that I need to separate in Excel that is in a single cell. The "Text to Columns" feature works great except for one snag.

In a single cell, I have `First Name`, `Last Name` & `Email address`. The last name and email addresses do not have a space between them, but the color of the names are different than the email.

Example (all caps represent colored names RGB (1, 91, 167), lowercase is the email which is just standard black text):

```
JOHN DOEjohndoe@acmerockets.com

```

So I need to put a space after DOE so that it reads:

```
JOHN DOE johndoe@acmerockets.com

```

I have about 20k rows to go through so any tips would be appreciated. I just need to get a space or something in between that last name and email so I can use the "Text to Columns" feature and split those up.

----

You can knock this out pretty quickly taking advantage of a how `Font` returns the `Color` for a set of characters that do not have the same color: it returns `Null`! Knowing this, you can iterate through the characters 2 at a time and find the first spot where it throws `Null`. You now know that the color shift is there and can spit out the pieces using `Mid`.

**Code** makes use of this behavior and `IsNull` to iterate through a fixed `Range`. Define the `Range` however you want to get the cells. By default it spits them out in the neighboring two columns with `Offset`.

```
Sub FindChangeInColor()

    Dim rng_cell As Range
    Dim i As Integer

    For Each rng_cell In Range("B2:B4")
        For i = 1 To Len(rng_cell.Text) - 1
            If IsNull(rng_cell.Characters(i, 2).Font.Color) Then
                rng_cell.Offset(0, 1) = Mid(rng_cell, 1, i)
                rng_cell.Offset(0, 2) = Mid(rng_cell, i + 1)
            End If
        Next
    Next
End Sub

```

**Picture of ranges and results**

![ranges and results](https://i.stack.imgur.com/UAygx.png)

The nice thing about this approach is that the actual colors involved don't matter. You also don't have to manually search for a switch, although that would have been the next step.

Also your neighboring cells will be blank if no color change was found, so it's decently robust against bad inputs.

**Edit** adds ability to change original string if you want that instead:

```
Sub FindChangeInColorAndAddChar()

    Dim rng_cell As Range
    Dim i As Integer

    For Each rng_cell In Range("B2:B4")
        For i = 1 To Len(rng_cell.Text) - 1
            If IsNull(rng_cell.Characters(i, 2).Font.Color) Then
                rng_cell = Mid(rng_cell, 1, i) & "|" & Mid(rng_cell, i + 1)
            End If
        Next
    Next
End Sub

```

**Picture of results again** use same input as above.

![edit results](https://i.stack.imgur.com/Ku9Uu.png)
