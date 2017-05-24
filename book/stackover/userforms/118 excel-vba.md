# SO item 118
I am using the below code to enable a user to drag and drop a label.

The code works fine - but I am looking for a way of 1)simplifying the code when dealing with several labels and 2)give the user the option to create a new label which has the same properties i.e drag/drop.

As it stands, the code specifically refers to specific labels i.e Label1 etc, I have to copy the code again and again to refer to all the labels I want (50+)

So essentially is there a way of having my code to automatically work for all labels, both existing and newly created?

```
Private x_offset%, y_offset%
Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As    Integer, _
ByVal X As Single, ByVal Y As Single)

If Button = XlMouseButton.xlPrimaryButton Then
 x_offset = X
 y_offset = Y 
End If

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)

If Button = XlMouseButton.xlPrimaryButton Then
Label1.Left = Label1.Left + X - x_offset
Label1.Top = Label1.Top + Y - y_offset
End If

End Sub

```

Thank you

----

A little late on the response, but here is how this is done. The idea is that you need to create a class module which can handle the events for the `Label`. Once you have the class in place to handle the event, you need to wire up the new/existing `Labels` to go through the class. This is commonly done by creating a `Collection` which holds all your class objects. Other than that, you just need to create a class object for each label (new or existing). The following pieces are needed:

*   UserForm1 with its code behind
*   LabelHolder class module

**LabelHolder** class module contains the code for an ideal "Label Holder". This is a simple class which holds a reference to a `MSForms.Label` and handles each one's events. Note that I have called the object `Label1` so that I could lazily copy your code. This `Label1` has **nothing to do** with the `Label1` on the `UserForm`; they have different scopes and are independent.

```
'class module code
Public WithEvents Label1 As MSForms.Label
Private x_offset%, y_offset%

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)

    If Button = XlMouseButton.xlPrimaryButton Then
        x_offset = X
        y_offset = Y
    End If

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)

    If Button = XlMouseButton.xlPrimaryButton Then
        Label1.Left = Label1.Left + X - x_offset
        Label1.Top = Label1.Top + Y - y_offset
    End If

End Sub

```

**UserForm1 code behind** shows the event for the button which creates a new `Label` and adds it to the `Collection`. It also stores the `Collection` which ensures that the class objects have a global scope and are not garbage collected early. There is also an `Initialize` event which shows how to add an existing `Label` to the fold.

```
'UserForm1 code behind
Dim labels As Collection

Private Sub CommandButton1_Click()

    If labels Is Nothing Then
        Set labels = New Collection
    End If

    Dim lbl As MSForms.Label
    Set lbl = Frame1.Controls.Add("Forms.Label.1")

    lbl.Caption = "testing"

    Dim holder As New LabelHolder
    Set holder.Label1 = lbl

    labels.Add holder

End Sub

Private Sub UserForm_Initialize()

    If labels Is Nothing Then
        Set labels = New Collection
    End If

    Dim holder As New LabelHolder
    Set holder.Label1 = Label1

    labels.Add holder

End Sub

```

Finally here is an image of the `UserForm1` which has default names for all the controls.

![image of user form](https://i.stack.imgur.com/50IWt.png)

Same form after clicking the button and dragging things around:

![picture after some work](https://i.stack.imgur.com/a3rqM.png)

All of this code shows how to connect a class module to dynamically created and original components on the User Form. It does not address how to create a new `Label` with the drag/drop, but it is possible. You would put that code in the class module and ensure that you have enough references back to the User Form in order to access the properties you need there.
