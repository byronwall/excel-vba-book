## understanding the Workbook Object Model

The Workbook is an object that serves two main purposes:

* Provide a foothold to a number of other more useful functions (e.g. Sheets)
* Provide a reference to the underlying data within a spreadsheet while working through a workflow

For the first point, the goal of the Workbook is to actually use its properties to do some task.  For the latter point, the Workbook is simply a container that holds data which is necessary to interact with while creating a workflow.  To be honest, there is very little of use within the Workbook object that is not a refernece to some other object.  Typically, the main tasks to actually be done with the Workbook are to Open, Save, and Close them.  That is, you move away from the Workbook object as quickly as you can because you just need a reference.
