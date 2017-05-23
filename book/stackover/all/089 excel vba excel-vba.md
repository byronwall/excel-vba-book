# SO item 089
This is the first time I have ever really worked with VBA stuff (the code I found is VBA right? :P ) so please forgive my inability to just...understand it by looking at it. :) I have experience with macros and know /some/ of that code is compatible, but I don't know how to write it...only know how to record it then use the bits I know work.

The Puzzle: I have a column which will adjust the weighted values of the row it is in (either makes the values slightly more or less, or leaves them alone), but the weights are entirely subjective so only need to be changed manually usually.

The issue is when needing to change a lot of them at once because it is tons of clicking each cell and changing the value one at a time by "guessing" which weights we need by entering numbers until it "looks good" (like I said subjective), so I found this:

[http://www.mrexcel.com/forum/excel-questions/51933-changing-cell-value-clicking-cell.html](http://www.mrexcel.com/forum/excel-questions/51933-changing-cell-value-clicking-cell.html)

Made some tiny changes to the version at the bottom of the page and have this right now:

```
Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
  Application.EnableEvents = False
  If Target.Cells.Count = 1 Then
    If Not Intersect(Target, Range("Q3:Q500")) Is Nothing Then
      Select Case Target.Value
      Case ""
        Target.Value = "+"
      Case "+"
        Target.Value = "-"
      Case "-"
        Target.Value = ""
      Case Else
        Target.Value = "+"
      End Select
        ActiveCell.Offset(0, 2).Range("A1").Select
    End If
  End If
  Application.EnableEvents = True
End Sub

```

Almost perfect, except the range of numbers will be slightly larger than "", +, or -. In fact the range will likely be 0.5 to 1.5 (50% to 150%).

It would be best if they could go up or down. I guess I could just create a huge list of replacements so instead of "", +, or - it would just be the numbers (Case 0.5 > Case 0.6 > Case 0.7 >>> Case 1 > Case 1.1 >>> Etc.) but obviously that would probably be a huge nightmare and /less/ efficient when having to click the cell 10 times to cycle through the numbers.

So there are two things I need to make it work perfectly:

1.  For it to change the NEXT cell over, either to the left or right by one.
2.  Instead of simple replacements, is it possible for it to do some math?

The idea is that clicking in column P to the left of Q will decrease the weight (by say 0.1), and by clicking in column R to the right of Q it will increase the weight slightly. And if they reach the bottom/top it will go to the other end (0.5 decreased will take the value up to 1.5, vice versa).

I tried tinkering a little with what I know about macros and tried adding the "ActiveCell.Offset(0, 1).Range("A1").Select" bit here and there, which will be nice for the end of the code so I can choose where the cursor is left afterwards...but putting it anywhere earlier...I see the active cell move to P or R before changes are made, but it still changes the original column Q.

I recognize I will need two copies of the code, one for Column R (decreasing) and one for Column P (increasing) but so long as I have at least one side I should be able to make the changes to mirror it. Also I know Data Validation might seem like a simple solution, but having an obscene amount of drop down arrows covering most of the sheet will make it illegible...

Thank you so much to anyone spending any time on helping to solve this puzzle, Corinne :)

----

I think I understand your question, if this doesn't do it, please add a picture and we'll get there.

Changing the `Value` based on `Value` is straightforward.

**Code** shows how to handle clicks (or arrow movements) into columns `P` or `R`. Going into `P` will decrease the value (with the wraparound like you mention). Going into `R` will go up.

```
Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
  Application.EnableEvents = False
  If Target.Cells.Count = 1 Then
    If Not Intersect(Target, Range("P3:P500")) Is Nothing Then

        If Target.Offset(0, 1).Value > 0.5 Then
            Target.Offset(0, 1).Value = Target.Offset(0, 1).Value - 0.1
        Else
            Target.Offset(0, 1).Value = 1.5
        End If

        Target.Offset(0, 1).Select
    End If

    If Not Intersect(Target, Range("R3:R500")) Is Nothing Then

        If Target.Offset(0, -1).Value < 1.5 Then
            Target.Offset(0, -1).Value = Target.Offset(0, -1).Value + 0.1
        Else
            Target.Offset(0, -1).Value = 0.5
        End If

        Target.Offset(0, -1).Select
    End If

  End If
  Application.EnableEvents = True
End Sub

```

The main idea is to check the value in column `Q` and add or subtract or wrap around depending on column and value. This code can be simplified and cleaned up, but it works which may be all you need.

_I am assuming that you have a number in column `Q` and are trying to increase or decrease that? If not, please add more detail._

**The nice part of this code is that it puts the selection back into column `Q`. This means that you can repeatedly hit the right or left arrow to increase or decrease, respectively**. This works because the selection changes but then is immediately put back in the center. You can type fast and increment or decrement as needed.

**Picture** shows the assumption of data in column `Q`. I used the code a couple times to change values to what is shown.

![results](https://i.stack.imgur.com/9T6BT.png)
