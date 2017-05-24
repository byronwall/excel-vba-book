# SO item 047
I want to run a macro that allows for the following steps:

1.  User clicks on shape A and runs macro
2.  Macro will record position and size properties of shape A

3.  User clicks on Shape B on a different slide

4.  Macro applies position and size properties of shape A to shape B
5.  User clicks on Shape C on a different slide
6.  Macro applies position and size properties of shape A to shape C etc...

So far I have been able to get the initial shape (Shape A's) properties, but am not sure how to let the user select the next shapes.

```
Dim w As Double
Dim h As Double
Dim l As Double
Dim t As Double

With ActiveWindow.Selection.ShapeRange(1)
    w = .Width
    h = .Height
    l = .Left
    t = .Top
End With

```

Appreciate the help!

* * *

See below for answer. If you haven't used forms before (like myself), the code that begins with "Private Sub CommandButton1_Click()" should NOT be inserted in the same module. Go to Insert > Userform, then drag two command buttons onto the UI box, and another "Userform code" window should appear. That new window is where the "Private Sub CommandButton1_Click()" code should go.

----

I think you will have trouble using click events for this. I would recommend creating macros and storing them on the Quick Access Toolbar. Once there, the keyboard shortcut is ALT+SOME NUMBER which can be quickly used.

For the code, the general idea is that you create the variables with `global` scope. This allows them to retain their values after the `Sub` finishing execution.

In the code below, `StoreDetails` will save, and `OutputDetails` will apply to newly selected object. The saved info will stay there so you can go from A to save and then apply to B, C, D without seeing A again.

**Code inside Module1**

```
Dim w As Double
Dim h As Double
Dim l As Double
Dim t As Double

Sub StoreDetails()
    With ActiveWindow.Selection.ShapeRange(1)
        w = .Width
        h = .Height
        l = .Left
        t = .Top
    End With
End Sub

Sub OutputDetails()
    With ActiveWindow.Selection.ShapeRange(1)
        .Width = w
        .Height = h
        .Left = l
        .Top = t
    End With
End Sub

```

Here is an article about [assigning macros to the Quick Access Toolbar](https://support.office.com/en-in/article/assign-a-macro-to-a-button-728c83ec-61d0-40bd-b6ba-927f84eb5d2c) if you need help there.
