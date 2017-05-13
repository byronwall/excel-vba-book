# overview of 4 adv processing

Advanced processing should include some recipe type sections that go through the more advanced aspects of working up VBA code.

This could focus on:

* Speed improvements and how to do it (disable screen, events, calculation) and how to undo it
* Working with arrays of values instead of outputting a cell at a time
* Cranking through an entire automated workflow without user interaction: creating new workbooks, worksheets, charts, formulas and then outputting it all to PDF
* Focus on the interplay of manual steps and code (sometimes you have to run part of the code to see what to do next; other times you can sit down and type the whole thing out)
* Cleaning up macro recorder code (some discussion about what works well/doesn't)
* How to avoid `Select` and why
* Using `DoEvents` to wait a set amount of time
* Using `Application.OnWait` (?) to do some thing at a regular time
* Parsing through existing formulas or values and manipulating with confidence
* Reading and writing to external files
* Working with the file system to do some processing
* Running through a folder or batch of files and doing something with each one
* Structuring code in a way that the different pieces can be called on their own
* Going through a workflow that involves using other office products
* Strategy for identifying cells using Styles and working through them; effectively a tag feature

The long list of sections here says that maybe there is enough code to put together a couple of "case study" type things that break down the development of an entire workflow.  This could be something from TDA related to charting/processing or some other thing.
