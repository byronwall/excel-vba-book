## comments on knowing Excel before starting VBA

The Basics of Excel chapter may not actually be needed. The problem with working through VBA is that it requires a baseline understanding of how to use Excel. In general, it requires more than a baseline understanding of how to use Excel -- it requires an advanced understanding of how Excel works. In particular, it's important to know what happens by default when you do something in Excel. This then makes it much easier to reason through how the VBA is going to function.

TODO: add explanations for other reasons why: using formulas, intuition, knowing what's possible.

An example: if you have never used the keyboard shortcut CTRL+arrow to jump around a block of data, you would never think to look for the Range.End() command which replicates this behavior. Instead, you might be tempted to iterate through every cell calling Offset(1) until the cell is empty.

A second example: most people are familiar with the AutoFill behavior by dragging the corner of the current cell. This is great most of the time but has the bad habit of trying to predict a series when you just want a constant. There is also the Fill command (and a keyboard shortcut `CTRl+D` for fill down and `CTRL+R` for right) that will copy the formula or the value without trying to predict the next value. Fill is much more useful when working with VBA since it's unlikely to secretly ruin your data.

Knowing that there are multiple ways to do something and knowing the quirks of specific commands is invaluable when working with Excel through VBA. You will have a much better intuition for what will happen if you know how Excel normally does things.

If you don't have this intuition, you can still learn VBA and be effective, but you may find yourself falling back on common programming techniques that are not "idiomatic" Excel.

This section might include some the VBA equivalent of common Excel techniques?

TODO: decide if this section is needed
