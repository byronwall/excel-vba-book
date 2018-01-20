### working with `Ranges` via advanced techniques

* Use the Offset-Intersect technique to get a block of data without its header
* Use the `AutoFilter` to filter a data set and then get the visible cells with `SpecialCells()`
* Use one of the techniques above to get a `Range` on one `Worksheet`; grab the corresponding `Range` on a another `Worksheet` to do some processing

#### Offset-Intersect

The Offset-Intersect is one of the most useful and simple approaches to creating a Range. The idea is that by using Intersect, you will avoid ever creating a Range that is bigger than some starting point.  THis means that you will not be able ot accidetanlly add a blank or neighboring column to your Range.  KNowing this, you can then take whatever steps are necessary to "remove" bad sections fromy our Range.  This is most comonly used to remove a header row from teh top of a Range.  If you are using Offset, the only rule is that you must make a valid move before calling INtersect.  To remove a header, assuming you have a range which is a block of data iwth headers, simply do: `Set rng = Intersect(rng, rng.Offset(1))`.  This gives you a new Range which has all of the cells of the first oen except for the first row.

TODO: add an image of how this works

Intersect used in this fashion is incredibly pwoerful.  You can do all sorts of wacky steps to filter out a Range and then Intersect against the orignal Range to ensure that you have not accidetanlly stepped outside your starting box.

#### AutoFilter and then SpecialCells

This approach is straight forward adn mirrors a cmomon operation in non VBA Excel.  You use an AutoFilter to filter out specific cells adn tehn you can select ony the visible cells.  In Excel, you can use `ALT+SEMICOLON` to only select visible cells.  Often times, you will not need to actually do this since Excel tries to help you when dealing with Hidden rows and colujmns.  Tpyically Excel will nto apply formatting to hidden cells and will also not fill a fomrula through them (assuming you used the Fill command).

In VBA, things are often more difficult because you are working with the underlying Range indepednet of whether or not the cells are hidden. To get around this, Excel provides the SpecialCells function which allows you to select a subset of cells based on some crtieria. wehn using the AutoFilter, the most common criteriojn to use is that of visibilty.  You can call `Range.SpecialCells(xlCellTypeVisisble)` to obtain a new Range which only coantians visible cells.

If you haev ever written a loop which does a `If rng.Hidden = True Then...` then you will be grateful to know that Excle VBA provides this feature automtically.  SpecialCells really is one of the most powerful ways to access Ranges in an intuitive fahsion that matches normal Excel.

#### The Duplicated Range on another Sheet

If you are working with multiple sheets that are the same, similar, or related, you will often find yourself using inforamtiom about one sheet to build a Range on another or several others.  THe probelm with Ranges however is that they are not allowed to span multiple Worksheets.  This means that if you want to apply some actiojn ot each `A1:A10` Range on each Worksheet, you will need to do it iteratively.  This can be a pain however if you built your Range usign code and not a direct address.  To get aroudn this, you can use the `Range.Address()` function to obtain an address for the Range.  The trick here is to use the `Address` function without parameters which will give you the local address without a Worksheet name. Y ou can then use that address on each of the other Worksheets, you access the given cells on that Worksheet.

This is anice way to replicate the functioanlity of Excel where you can sleect multiple Worksheets with CTRL or SHIFT and then apply some acitojn to all of them.  The realyl nice thing about VBA however is that you can apply an action that is aware of the Worksheet on which it is acting.  This si quite nice becasue the normal muti edit feature do the exact same steps to all spredahseets whereas you may want to use `End` or soemthign in your code.
