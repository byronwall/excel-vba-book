# SO item 010
I have a Word document that is "form-fillable", i.e. it has content control objects in it such as rich text and date picker content controls. I am looking to extract the data from specific fields into Excel. For example, every form has the project title, start date, and manager. I would like 1 row for that form with those three pieces of data. Eventually this will need to be done for a few hundred of these forms every few months, but for now I'd like to just start with one.

I envision having a button on the Excel sheet that will run VBA code to pull the data from the Word document, and populate the proper cells in the sheet. With the filepath for the Word doc being specified by the user.

I am new to VBA. How do I point my code at the right file, and then again at the specific field I need? Do I give the fields titles in the Word doc?

This is in MS Office '13

----

Your application is going to have a large number of specific details which are difficult to address without the specifics. To get started, here is some very simple code to get the Date from a drop-down in Excel from Word.

Note to work with Word, you need to create a reference to "Microsoft Word 15.0 Object Library" on the Excel side. Tools -> References in the VBA Editor.

![reference addition](https://i.stack.imgur.com/zEKHX.png)

When working across applications, it can help to write the VBA in Word and then move it over to Excel once you have the properties you want.

This VBA for the Excel side of the equation:

```
Sub GetDataFromWordFile()

    Dim wrd As Word.Application
    Dim file As Word.Document

    Set wrd = New Word.Application
    Set file = wrd.Documents.Open("test.docx")

    Range("A1") = file.ContentControls(1).Range.Text

    file.Close
    wrd.Quit

End Sub

```

My Word file includes a single ContentControl. Not sure how you address the ones you want (trial and error?).

![Word file](https://i.stack.imgur.com/PPxAc.png)

The Excel side drops the date in the right place.

![Excel side](https://i.stack.imgur.com/LpyAO.png)
