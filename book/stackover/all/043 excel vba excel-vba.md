# SO item 043
I have an excel file which I import data from bloomberg via the function BDH. I can easily update bloomberg formulas, but the other columns with excel formulas are not updated, so currently I need to drag down the excel formulas everyday. I have already tried to use the code but it does not work. Can someone help me on that? Thank you very much

```
Sub update_formulas()

Activeworkbook.RefreshAll

End sub

```

----

CTRL+ALT+F9 is the keyboard shortcut for a full recalculation.

`Application.CalculateFullRebuild` is another way to force a refresh of an entire workbook's formulas if want to use VBA.

Note that `RefreshAll` is only for refreshing `Data` related items. It is the same as going to `Data->Refresh->Refresh All`. It will update Pivot Tables and external connections. It will generally not update formulas unless they are referencing the data / Pivot Table that was updated.
