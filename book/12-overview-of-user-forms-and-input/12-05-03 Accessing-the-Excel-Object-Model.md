### Accessing the Excel Object Model

From a UserForm, you have full access to the Excel Object Model. This can be very handy if you are trying to access information from the USerForm to determine what information to show in the Form. It can also be helpful if oyu want to make changes to the underlying spreadsheet from a USerForm without leaving the form. Both of those options are very common and very easy to do with UserForms. In general, any code that can run without a USerForm present can be run with a USerForm. There are some limiations when it comes to the user's ability to Select items with a From visible, but you are not limited in calling the same commands from VBA (TODO: is that right?). The exception ehre is that if the form is `ShowModal = False` then the user is able to make selections while the from is bisible.

There is no real limit to what you can do from a SuerForm. A couple of examples to give you a feel:

- present a list of all open Workbooks so that they user can select which one that want to process
- Create a form that can process all of the selected CHarts.
- Present a ListBox with the unique values from all of the AutoFilters that are active. Allow the user to selectively remove or chagne those filters without having to use the normal drop downs.
