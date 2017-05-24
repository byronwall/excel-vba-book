# SO item 104
I am writing a VBA application in Excel. I have a Userform that dynamically builds itself based upon the data contained in one of the worksheets. All of the code that creates the various comboboxes, textboxes and labels is working. I created a class module to trap OnChange events for the Comboboxes, and again this works as expected. Now I have a need to trap OnChange events for some of the textboxes, so I created a new class module modelled on that for the comboboxes to trap the events.

```
Public WithEvents tbx As MSForms.TextBox

Sub SetTextBox(ctl As MSForms.TextBox)
Set tbx = ctl
End Sub

Public Sub tbx_Change()
Dim LblName As String

MsgBox "You clicked on " & tbx.Name, vbOKOnly

End Sub

```

The message box is just so that I can confirm it works before I go further. The problem I'm getting is in the UserForm code module:

```
Dim TBox As TextBox
Dim tbx As c_TextBoxes

'[...]

Set TBox = lbl
Set tbx = New c_TextBoxes
tbx.SetTextBox lbl
pTextBoxes.Add tbx

```

This throws up a type mismatch error at `Set TBox=lbl`. It's the **exact** same code that works fine for the ComboBox, just with the variables given approriate names. I've stared at this for 2 hours. Anyone got any ideas? Thanks for any pointers.

Edit - Here's the full userform module that I'm having trouble with:

```
Private Sub AddLines(FrameName As String, PageName As String)
Dim Counter As Integer, Column As Integer
Dim obj As Object
Dim CBox As ComboBox
Dim cbx As c_ComboBox
Dim TBox As TextBox
Dim tbx As c_TextBoxes
Dim lbl As Control

Set obj = Me.MultiPage1.Pages(PageName).Controls(FrameName)
If pComboBoxes Is Nothing Then Set pComboBoxes = New Collection
If pTextBoxes Is Nothing Then Set pTextBoxes = New Collection

For Counter = LBound(Vehicles) To UBound(Vehicles)
     For Column = 1 To 8
     Select Case Column
     Case 1
         Set lbl = obj.Add("Forms.Label.1", "LblMachine" & FrameName & Counter, True)
    Case 2
        Set lbl = obj.Add("Forms.Label.1", "LblFleetNo" & FrameName & Counter, True)
    Case 3
        Set lbl = obj.Add("Forms.Label.1", "LblRate" & FrameName & Counter, True)
    Case 4
        Set lbl = obj.Add("Forms.Label.1", "LblUnit" & FrameName & Counter, True)
    Case 5
        Set lbl = obj.Add("Forms.ComboBox.1", "CBDriver" & FrameName & Counter, True)
    Case 6
        Set lbl = obj.Add("Forms.Label.1", "LblDriverRate" & FrameName & Counter, True)
    Case 7
        Set lbltbx = obj.Add("Forms.TextBox.1", "TBBookHours" & FrameName & Counter, True)
    Case 8
        Set lbl = obj.Add("Forms.Label.1", "LblCost" & FrameName & Counter, True)
    End Select
    With lbl
        Select Case Column
        Case 1
            .Left = 1
            .Width = 60
            .Top = 10 + (Counter) * 20
            .Caption = Vehicles(Counter).VType
        Case 2
            .Left = 65
            .Width = 40
            .Top = 10 + (Counter) * 20
            .Caption = Vehicles(Counter).VFleetNo
        Case 3
            .Left = 119
            .Width = 50
            .Top = 10 + (Counter) * 20
            .Caption = Vehicles(Counter).VRate
        Case 4
            .Left = 163
            .Width = 30
            .Top = 10 + (Counter) * 20
            .Caption = Vehicles(Counter).VUnit
        Case 5
            .Left = 197
            .Width = 130
            .Top = 10 + (Counter) * 20
            Set CBox = lbl 'WORKS OK
            Call CBDriver_Fill(Counter, CBox)
            Set cbx = New c_ComboBox
            cbx.SetCombobox CBox
            pComboBoxes.Add cbx
        Case 6
            .Left = 331
            .Width = 30
            .Top = 10 + (Counter) * 20
        Case 7
            .Left = 365
            .Width = 30
            .Top = 10 + (Counter) * 20
            Set TBox = lbl 'Results in Type Mismatch
            Set tbx = New c_TextBoxes
            tbx.SetTextBox TBox
            pTextBoxes.Add tbx
        Case 8
            .Left = 400
            .Width = 30
            .Top = 10 + (Counter) * 20
        End Select
    End With
    Next
Next
obj.ScrollHeight = (Counter * 20) + 20
obj.ScrollBars = 2

End Sub

```

And here's the c_Combobox class module:

```
Public WithEvents cbx As MSForms.ComboBox

Sub SetCombobox(ctl As MSForms.ComboBox)
    Set cbx = ctl
End Sub

Public Sub cbx_Change()
Dim LblName As String
Dim LblDriverRate As Control
Dim i As Integer

    'MsgBox "You clicked on " & cbx.Name, vbOKOnly
    LblName = "LblDriverRate" & Right(cbx.Name, Len(cbx.Name) - 8)
    'MsgBox "This is " & LblName, vbOKOnly

    'Set obj = Me.MultiPage1.Pages(PageName).Controls(FrameName)
    Set LblDriverRate = UFBookMachines.Controls(LblName)
    For i = LBound(Drivers) To UBound(Drivers)
        If Drivers(i).Name = cbx.Value Then LblDriverRate.Caption = Drivers(i).Rate
    Next
End Sub

```

And finally, here's the c_TextBoxes class module:

```
Public WithEvents tbx As MSForms.TextBox

Sub SetTextBox(ctl As MSForms.TextBox)
    Set tbx = ctl
End Sub

Public Sub tbx_Change()
Dim LblName As String
    'Does nothing useful yet, message box for testing
    MsgBox "You clicked on " & tbx.Name, vbOKOnly

End Sub

```

----

After some quick testing, I am able to reproduce your error if I declare `TBox as TextBox`. I do not get an error if I declare `TBox as MSForms.TextBox`. I would recommend declaring all your `TextBox` variables with the `MSForms` qualifier.

**Test code** is situated similar to yours. I have a `MultiPage` with a `Frame` where I am adding a `Control`.

```
Private Sub CommandButton1_Click()

    Dim obj As Object
    Set obj = Me.MultiPage1.Pages(0).Controls("Frame1")

    Dim lbl As Control
    Set lbl = obj.Add("Forms.TextBox.1", "txt", True)

    If TypeOf lbl Is TextBox Then
        Debug.Print "textbox found1" 'does not execute
    End If

    If TypeOf lbl Is MSForms.TextBox Then
        Debug.Print "textbox found2"

        Dim txt1 As MSForms.TextBox
        Set txt1 = lbl 'no error
    End If

    If TypeOf lbl Is MSForms.TextBox Then
        Debug.Print "textbox found3"

        Dim txt As TextBox
        Set txt = lbl 'throws an error
    End If

End Sub

```

I am not sure why the qualifier is needed for `TextBox` and not `ComboBox`. As you can see above, a good test for this is the `If TypeOf ... Is ... Then` to test which objects are which types. I included the first block to show that `lbl` is not a "bare" `TextBox`, but, again, I have no idea why that is. Maybe there is another type of `TextBox` out there that overrides the default declaration?
