## print layout and exporting

THis section will focus on the print and export related details of a Worksheet. In particular, it will focus on the details that are typically accessed through the Page Layout menu. This is one of the unique aspects of Worksheets because they are the holder of the print/export information. The main details related to this are:

- Print area
- Page layout -- this is a very large object with a lot of properties to be set
- Exporting and printing

The details in this section can be a real time saver because one of the more tedious aspects of working with Excel is ensuring that your reports/graphs/data will print or export correctly. Being able to control these properties with VBA makes it possible to quickly apply the same formatting to a large number of Worksheets without having to click nine million times.

When editing the Page Layout, you can change nearly everything. The one thing ot be aware of is related to printers. There are a number of settings in the Worksheet that are internally tied to the defualt (or active) printer. This shows up if you are attmepting to set the page size specifically. IF you always use the same printer or have coworkers who use the smae printers, you amy not notice these pissues. It becomes a serious probelm when you ar etrying ot make code work for multiple differnet printers that support or dientify page sizes differently.

The best way to see what is savialable for page settings is to record a macro and change one thing. Excel is a bit agressive at including all possible settings that coudl ahve hcanged. This is very nice if you want to grab some setings nad work them into your code.

There are a couple of other items to describe so you know what they are:

- Using `Zoom` and `FitToPages` to set the number of pages that the output will be included in (TODO: review)
- TODO: add others

Also, be aware that hcanging the print settings is a per Worksheet change. This amy be ovvious since hte peroperies are off the Worksheet, but it is easy to forget this. The nice thing however s that you can just iterate your WOrksheets and apply the same settings to all of them. THis is oen of the greatest time savers comapred to changing properties in Excel (TODO: can these be changed with multi selecteion?).
