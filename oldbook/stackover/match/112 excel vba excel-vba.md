# SO item 112
I'm trying to assign a cell a formula using VBA, but every time I run the code it assign to the cell the result not the formula itself, like when I enter the formula using the excel spreadsheet.

Does anyone know how to display the formula using a macro within a cell and not the formula result.

I'm asking this because I need to insert this line within an existing sheet among other data and then run another macro to keep it updated, and the macro depends on this formula.

The code I'm using

```
Cells(C, 9).Formula = Application.Index(Plan2.Range("B2:D10000"), _
Application.Match(Plan1.Range("B" & C) & Range("F" & C), Plan2.Range("A2:A10000"), 0), 3)

```

As you can see this formula depends on the row that its inserted.

----

Your `Formula` is ultimately being reduced to whatever is returned by `Application.Index`. Unless the value being returned there is an actual formula string then you will just get a number as the result and this is set to the `.Formula`.

If you want to actually set the formula, you need to create a string in VBA that represents the formula to use. In this case, that string would look something like:

```
Cells(C, 9).Formula = "=INDEX(Plan2!B2:D10000, MATCH(Plan1!B" & C & "..."

```

where you concatenate in the dynamic parts. The end result needs to look like a normal formula. The `Application.XXX` and `Application.WorksheetFunction.XXX` functions return actual results, not pieces that can be combined to create a formula.
