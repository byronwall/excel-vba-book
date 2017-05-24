# SO item 099
I created a Userform in Word which imports 3 columns of data from an excel sheet, inserts it into bookmarks and in the name of the word document and saves it as a pdf.

Now I wanted to add a Listbox into the form to be able to add, modify and delete the inputs manually which are usually imported from the excel sheet .

I already figured out how to add data from 3 textboxes into a 3 Column Listbox but even after googling for hours I can't find a good way to modify selected data.

VB.net has the .selectedItem property, VBA does not. Can anybody give me an example how to modify a multi column listbox with the values of 3 textboxes?

Thanks in advance

----

You need to iterate through `ListBox.Selected` and check if it is `True`. Once you get a `True` one, you can process that item.

**Sample code** adds some items with columns and sets up a click event to run through the `Selected` items.

```
Private Sub ListBox1_Click()
    Dim i As Integer
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            Debug.Print ListBox1.List(i, 0)
        End If
    Next i
End Sub

Private Sub UserForm_Initialize()
    ListBox1.AddItem "test"
    ListBox1.AddItem "test1"
    ListBox1.AddItem "test2"

    ListBox1.ColumnCount = 3
    ListBox1.ColumnHeads = True

    ListBox1.List(1, 0) = "change to 1,0"
    ListBox1.List(1, 1) = "change to 1,1"
    ListBox1.List(1, 2) = "change to 1,2"
End Sub

```

**Picture of form with Immediate window** after clicking each item in turn.

![UserForm](https://i.stack.imgur.com/1hLwe.png)
