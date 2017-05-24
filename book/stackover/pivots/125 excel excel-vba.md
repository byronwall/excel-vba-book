# SO item 125
I have the below code that based on what's entered in K2 it filters my pivot table based on that. I keep getting an error with a line that says (Set Field = pt.PivotFields) which calls the name of the field. The field I'm trying to influence is located in field name Locations/ sub field loc name. When I recorded myself changing it the code call the field the following:

```
ActiveSheet.PivotTables("PivotTable1").PivotFields( _
  "[Locations].[Loc Name].[Location Name]").VisibleItemsList = Array( _
  "[Locations].[Loc Name].[Location Name].&[CENTRAL MISSISSIPPI MED CTR  (CMS-1)]")

'ActiveSheet.PivotTables("PivotTable1").PivotFields( _
  "[Locations].[Loc Name].[Location Acronym]").VisibleItemsList = Array("")
'ActiveSheet.PivotTables("PivotTable1").PivotFields( _
  "[Locations].[Loc Name].[Location Number]").VisibleItemsList = Array("")

```

Any help on how what I put in cell K2 calls to the corresponding location in the pivot field will be great.

The Code:

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

If Intersect(Target, Range("K2:K3")) Is Nothing Then Exit Sub

Dim pt As PivotTable
Dim Field As PivotField
Dim NewCat As String

  Set pt = Worksheets("Fact Trans").PivotTables("PivotTable1")
  Set Field = pt.PivotFields("[Locations.Loc Name.Location Name]").VisibleItemsList
  NewCat = Worksheets("Fact Trans").Range("K2").Value

  With pt
    Field.ClearAllFilters
    Field.CurrentPage = NewCat
    pt.RefreshTable
  End With

End Sub

```

----

You are trying to assign an array of `String` to a `PivotField` which is not going to work. This is a result of adding `VisibleItemsList` to the end of the `PivotFields` call.

**Solutions**

*   If you really just want the `PivotField`, get rid of `VisibleItemsList`.
*   If what you really want is that list, you need to `Dim Field as Variant` so it can run. That would break downstream code though, so I doubt this is your goal.

I suspect you want the former which would be:

```
Set Field = pt.PivotFields("[Locations].[Loc Name].[Location Name]")

```

You can take a look at the [MS support for `VisibleItemsList`](https://msdn.microsoft.com/en-us/library/bb242470(v=office.12).aspx) if that's the property you really want to work with. Your macro probably included that in the code because you were changing a manual filter at the time (which is what that property controls).
