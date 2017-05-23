# SO item 075
I don't quite understand the differences between the combobox properties `SelText` and `Text`.

If I want to send the content of a combobox as a parameter to another procedure, should I send `.text` or `.selText`?

If I want to make enter text into a combobox using a macro, should I write the text in `.selText` or `.Text`?

----

The difference is really given in the name (**Sel**Text vs. Text) where **Sel** stands for `Selected`. One is used to return or modify the selected text (i.e. `SelText`) and the other is used to return or modify the entire text (i.e. `Text`).

If no text is selected in the ComboBox then they return and modify the same value.

I suspect you want to use `Text` unless you are specifically interested in the selected text.

This appears to be consistent for an ActiveX control on a `Worksheet` or for a control on a User Form.
