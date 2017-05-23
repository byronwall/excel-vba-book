# SO item 098
I have a workbook where I need to be able to click on a single cell of a worksheet and hit my command button. That copies and pastes the cell value to the first blank cell in column E on a different worksheet within the same workbook. When I just run the macro by itself, it works fine. But when I paste the code into a command button, it gives me a couple of runtime error 1004's. The most common error is "Select method of range class failed" and refers to the code line that tells it to select Range (E4). Here is the code:

```
Private Sub CommandButton1_Click()

' Choose player from Player list and  paste to Draft list.

    Sheets("Players").Select
    Selection.Select
    Selection.Copy

    Sheets("Draft").Select
    Range("E4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select
    Selection.PasteSpecial _ 
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub

```

----

**TL;DR**, couple options to resolve this, in order of preference:

1.  [Stop using `Select`](http://stackoverflow.com/q/10714251/4288101) to access cells
2.  Qualify your call to `Range("E4")` when executing code in a `Worksheet` object by using `Application.Range("E4")` or `Sheets("Draft").Range("E4")` or `ActiveSheet.Range("E4")`
3.  Move the code to `ThisWorkbook` or a code module and call that `Sub` from the event.

* * *

Here is the lengthy part that attempts to explain _why_ your code does not work.

This all comes down to: where is the code executing? Different execution contexts will behave differently when you use unqualified references to `Cells` `Range` and a number of other functions.

Your original code likely ran inside `ThisWorkbook`, a code module, or possibly in the code file for sheet `Draft`. Why do I guess this? Because in all of those places a call to `Range("E4")` would be acceptable to get the cell `E4` on sheet `Draft`. Cases:

*   `ThisWorkbook`and a code module will execute `Range` on the `ActiveSheet` which is `Draft` since you just called `Select` on it.
*   Inside `Draft` will execute `Range` in the context of `Draft` which is acceptable since that is the `ActiveSheet` and the place where you are trying to get cell `E4`.

Now what happens when we add an ActiveX `CommandButton` to the mix? Well that code is added to the `Worksheet` where it lives. This means that the code for the button can possibly execute in a different context than it did before. The only exception to this is if the button and code are both on sheet `Drafts`, which I assume not since you `Select` that sheet. For demonstrations, let's say the button is located on sheet `WHERE_THE_BUTTON_IS`.

Given that sheet, what is going on now? Your call to `Range` is now executed in the context of sheet `WHERE_THE_BUTTON_IS` **regardless** of the `ActiveSheet` or anything else you do outside of the call to `Range`. This is because the call to `Range` is _unqualified_. That is, there is no object to provide scope to the call so it runs in the current scope which is the `Worksheet`.

So now we have a call to `Range("E4")` in sheet `WHERE_THE_BUTTON_IS` which is trying to `Select` the cell. This is forbidden because sheet `Draft` is the `ActiveSheet` and

> **Thou shalt not `Select` a cell on a `Worksheet` that is not the `ActiveSheet`**

So with all of this, how do we resolve this issue? There are a couple of ways out:

1.  **[Stop using `Select` to manipulate cells](http://stackoverflow.com/q/10714251/4288101)**. This gets away from the main problem here, quoted above. This assumes your button lives on the same sheet as the `Selection` to copy/paste.

* * *

```
Private Sub CommandButton1_Click()

    Sheets("Draft").Range("E4").End(xlDown).Offset(1).Value = Selection.Value

End Sub

```

* * *

1.  **Qualify the call to `Range`** so that it executes in the proper context and chooses the right cell. You can use the `Sheets("Draft").Range` object to qualify this or `Application.Range` instead of the bare `Range`. I highly recommend option 1 instead of trying to figure out how to make `Select` work.

* * *

```
Private Sub CommandButton1_Click()
    Sheets("Players").Select
    Selection.Copy

    Sheets("Draft").Select

    'could also use Application.Range here
    Sheets("Draft").Range("E4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select

    Selection.PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

```

* * *

1.  Move the code back to a `Sub` that is outside of the `Worksheet` object and call it from the `CommandButton1_Click` event.
