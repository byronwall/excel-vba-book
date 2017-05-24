# SO item 045
I have the below snippet of code:

```
For i = 2 To n
    Postcode = Cells(i, 3)

    Cells(i, "M") = Postcode

    On Error Resume Next
    EndFrameOutput = Application.WorksheetFunction.VLookup(Postcode, Dealerws3.Range("C3:D" & LastRowDealer), 2, False)
    On Error GoTo 0

    Cells(i, 4) = EndFrameOutput
Next

```

The resulting output seems to fill the cells where there should be an N/A, with the previous successfully looked up value.

Eg: if i have this look up table:

```
Postcode    |   x
------------+-------
AB12 3BJ    |   1
IV1 1RY     |   2

```

And this Search Array:

```
Postcode
----------
AB12 3BJ
BE49 3GK
CG89 6KL
IV1 1RY
ML47 1KK

```

using my code, returning column 2 I get...

```
Postcode    |   Looked up Value
------------+-------------------
AB12 3BJ    |   1
BE49 3GK    |   1
CG89 6KL    |   1
IV1 1RY     |   2
ML47 1KK    |   2

```

instead of

```
Postcode    |   Looked up Value
------------+--------------------
AB12 3BJ    |   1
BE49 3GK    |   n/a
CG89 6KL    |   n/a
IV1 1RY     |   2
ML47 1KK    |   n/a 

```

Any ideas on how I can adapt my code?

Any help really appreciated!

thanks,

Colin

----

Use `Application.VLookup` instead of `Application.WorksheetFunction.VLookup`. The latter throws an application error when the query is not found. The former returns the error as the result which you can then deal with. Since you just want the `#N/A` as the output, you don't have to do anything special with it.

And get rid of those `On Error` calls. You won't need them with the different function, but in general you should avoid using them.

```
For i = 2 To n
    Postcode = Cells(i, 3)

    Cells(i, "M") = Postcode

    EndFrameOutput = Application.VLookup(Postcode, Dealerws3.Range("C3:D" & LastRowDealer), 2, False)

    Cells(i, 4) = EndFrameOutput
Next

```

Here is a good reference on the difference between those functions. [http://dailydoseofexcel.com/archives/2004/09/24/the-worksheetfunction-method/](http://dailydoseofexcel.com/archives/2004/09/24/the-worksheetfunction-method/)
