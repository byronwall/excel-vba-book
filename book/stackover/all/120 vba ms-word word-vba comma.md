# SO item 120
I am extracting text between parens in the active window title bar. That part is working great (thanks to some help I received here previously!). Now I want to create two separate macros - one that returns only the first name, and the other that returns only the last name.

My active window title bar looks something like this:

some text to the left (HENDERSON,TOM) some text to the right (there is no space after the comma)

The last name macro works perfectly. It looks like this:

```
Sub a1LastName()
    'Extract last name of patient from title bar (between parens)
    Dim strPatientName As String
    Dim OpenPosition As Integer '(open paren marker)
    Dim closeposition As Integer '(close paren marker)
    OpenPosition = InStr(ActiveDocument.ActiveWindow.Caption, "(")
    closeposition = InStr(ActiveDocument.ActiveWindow.Caption, ")")
    strPatientName = Mid(ActiveDocument.ActiveWindow.Caption, _
        OpenPosition + 1, closeposition - OpenPosition - 1)
    Dim c As Long
    c = InStr(strPatientName, ",")
    strPatientName = Left(strPatientName, c - 1)
    Selection.TypeText strPatientName
End Sub

```

The second macro is identical to the first, except that the second-to-last line of code has a "Right" instead of a "Left" instruction:

```
Sub a1FirstName()
    'Extract first name of patient from title bar (between parens)
    Dim strPatientName As String
    Dim OpenPosition As Integer '(open paren marker)
    Dim closeposition As Integer '(close paren marker)
    OpenPosition = InStr(ActiveDocument.ActiveWindow.Caption, "(")
    closeposition = InStr(ActiveDocument.ActiveWindow.Caption, ")")
    strPatientName = Mid(ActiveDocument.ActiveWindow.Caption, _
        OpenPosition + 1, closeposition - OpenPosition - 1)
    Dim c As Long
    c = InStr(strPatientName, ",")
    strPatientName = Right(strPatientName, c - 1)
    Selection.TypeText strPatientName
End Sub

```

Here's my problem: The "first name" macro always returns the last name minus the first four characters, followed by the first name, instead of simply the first name.

The only examples I'm able to find anywhere in Google land are specifically for Excel. I have combined through my VBA manuals, and they all give similar examples as I have used for extracting the text to the right of a character.

What am I doing wrong?

----

Seems like everyone has keyed in on `Split` to get the first/second part of the name. You can also use `Split` to get rid of the parentheses. This works if you know that you will only (and always) have a single `(` and `)`.

**Code** gives the main idea. You can use `Split` to get the part of the `String` which does not include the `(` or `)` and then do it again to get either side of the `,`.

```
Sub t()

    Dim str As String
    str = "(Wall,Byron)"

    Dim name_first As String
    Dim name_last As String

    'two splits total
    name_last = Split(Split(str, "(")(1), ",")(0)
    name_first = Split(Split(str, ")")(0), ",")(1)

    'three split option, same code to remove parentheses
    name_last = Split(Split(Split(str, "(")(1), ")")(0), ",")(0)
    name_first = Split(Split(Split(str, "(")(1), ")")(0), ",")(1)

End Sub

```

The code above presents a `two Split` and a `three Split` option. The main difference is that the three Split variety uses the same code to strip off the parentheses, and the only change is which side of the `,` to grab. The two Split options take advantage of the fact that splitting on a comma removes one the parentheses for free. The indices there are a little more complicated.
