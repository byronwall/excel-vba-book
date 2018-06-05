## limitations of UDfs

This section will focus on the aspects of UDFs where you are limited.  There are couple of key things to remember here:

* A UDF is not allowed to change the Workbook, Worksheet, or a Range -- no side effects are allowed
* A UDF will only update if the cells it refers to change
* You can mark a UDF as Volatile, but this may create other problems (namely speed)
* UDFs are allowed to use global variables but you can wreck this process by having errors while they execute
* UDFs inside an addin can pollute a spreadsheet that might be used by someone without that addin
* You can debug a UDF but not by using the Evaluate Formula option that might be familiar to more people
