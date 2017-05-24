# SO item 004
I got VBA code to populate some T-SQL query data in an Excel file. In that data, one column contains values of Red, Amber, Green and N/A. Now I want background Color for according to those values (Red, Amber, Green and White). How can I do this in VBA?

**Edited:** I need Like this:

```
id firstname lastname complaint
1  paul      nixon    RED
2  John      nathon   RED
3  sera      teag     AMBER
4  CLARE     walker   GREEN

```

Now I want background color for column 'complaint' according to cell value, like if cell value RED I want that background color also RED etc.. in VBA code.

----

Changing the background color of a cell is simple. Determining what color to change it to is the key step here. If you know that those 4 colors are the only options, I would just pound out the cases and set the colors. If you find this growing to more colors, you may want to define them in a Dictionary and do a lookup instead of the SELECT-CASE construction.

This simple code would work with your example. You will want to define the Range better (probably not "D2:D5") based on your real application and tweak the colors.

```
Sub ColorWithText()

    Dim cell As Range

    For Each cell In Range("D2:D5")
        Select Case UCase(cell.Value)
            Case "RED"
                cell.Interior.Color = RGB(255, 0, 0)
            Case "AMBER"
                cell.Interior.Color = RGB(255, 191, 0)
            Case "GREEN"
                cell.Interior.Color = RGB(0, 255, 0)
            Case "WHITE"
                cell.Interior.Color = RGB(255, 255, 255)
        End Select
    Next cell
End Sub

```

Here is a picture of my Excel instance after the code runs. ![image with colors applied](https://i.stack.imgur.com/BNyEa.png)
