### outputs

THe next item to hit are the otuputs of a workflow. Very often, the outputs are obviosu because you had some task to complete with VBA, and the outputs are simply the results of that task. Where things become more complicated is when you string together steps and the output of one becomes the input for the next. When that happens, you often have to decided what intermediate format is best for hte transfer. You may or may not settle on a format that is easily human consumable. There are tradeoffs here that will be discussed later. The output of a workflow can be a number of things:

* A string, number, cell, row, column, or table of data that was processed by the VBA
* A chart
* A collection of shapes
* A worksheet that includes any of the items above
* A workbook that includes a number of constructued worksheets
* A change to the fomratting of a number of cells
* A change to the properties of a Range, WOrksheet or Workbook
* A new text file written to disk
* Some result output to the Clipboard
* Pages of physical paper if your VBA prints
* Some change to the filesystem or disk
* Some other program opened or run with specific parameters

This is a shortened list since the possibilities here are closer to endless. The idea however is that you cna effect a large amount of change from VBA and so your possible outputs can be quite numerous. A tpyicaly workflow will accumulate a large number of these outoputs indviudally and will then produce some final product which highlights some of those outputs.
