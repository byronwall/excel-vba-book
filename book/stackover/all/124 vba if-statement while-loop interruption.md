# SO item 124
So the code I have below attempts to find WIP in column H. If we find WIP: copy 3 cells and make 10 replicas of them in the next column either in the same row or the next available row.

For some reason the code only runs the loop successfully for the first "WIP" value and then gives a code interruption error. Can someone see why this keeps happening?

Thank you, Ori

Sub Step1_update()

Dim dblSKU As Double Dim strDesc As String Dim strType As String Dim BrowFin As Integer Dim Browfin1 As Integer Dim Counter As Integer Dim Trowfin As Integer

Counter = 0

Worksheets("Final").Activate

Trowfin = 5 BrowFin = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

```
'loop 1
Do While Trowfin < BrowFin

    'If 1 (set the 3 values)
    If Range("H" & Trowfin).Value = Range("H3").Value Then

         dblSKU = Range("F" & Trowfin).Value
         strDesc = Range("G" & Trowfin).Value
         strType = Range("H" & Trowfin).Value

         'Find the last used row in Col J
         Browfin1 = (ActiveSheet.Range("J" & Rows.Count).End(xlUp).Row)

         Counter = 0
         'paste values 15 times
         Do While Counter < 15

            'If 2
            If Browfin1 > (Trowfin + Counter) Then

                   Range("J" & Browfin1).Value = dblSKU
                   Range("K" & Browfin1).Value = strDesc
                   Range("L" & Browfin1).Value = strType

            ElseIf Browfin1 < (Trowfin + Counter) Then

                   Range("J" & (Trowfin + Counter)).Value = dblSKU
                   Range("K" & (Trowfin + Counter)).Value = strDesc
                   Range("L" & (Trowfin + Counter)).Value = strType

            Else

                   Range("J" & (Trowfin + Counter)).Value = dblSKU
                   Range("K" & (Trowfin + Counter)).Value = strDesc
                   Range("L" & (Trowfin + Counter)).Value = strType

            End If

        'Loop to paste the WIP 15 times
        Loop

             Trowfin = Trowfin + 1
             Counter = 0

    'If cell (H...) is not a WIP
    Else

        Trowfin = Trowfin + 1

    'If 1
    End If

'loop 1
Loop

```

End Sub

----

Your previous code looked like:

```
  Else
    Range("J" & (Trowfin + Counter)).Value = dblSKU
    Range("K" & (Trowfin + Counter)).Value = strDesc
    Range("L" & (Trowfin + Counter)).Value = strType

    Counter = Counter + 1
  'If 2
  End If
Loop

```

You should do this instead:

```
  Else
    Range("J" & (Trowfin + Counter)).Value = dblSKU
    Range("K" & (Trowfin + Counter)).Value = strDesc
    Range("L" & (Trowfin + Counter)).Value = strType

  'If 2
  End If
  Counter = Counter + 1
Loop

```

Your current code has the increment inside the `Else` which means there is a chance that it will not be hit when the loop executes. If this happens, it your loop will go on infinitely, crashing Excel or causing the code interruption.

If you want to loop based on a counter, you need to ensure that the counter will reach the exit condition in a non-infinite length of time.
