### Helpful COmmands

There are a couple of commands that exist outside of addins that become far more useful inside the addin. They are included below for reference:

- `ThisWorkbook` refers to the workbook that contains the code being executed. This is the surefire way to refer to the XLAM file that is running isntead of the ActiveWorkbook. IN general, your addin will never be the ActiveWorkbook. This becomes relevant if your addin workbook contains sheets of data that may need to be acesssed during runtime. You woud use THisWorkbook to refer to those sheet.
- TODO: add any other commands that are addin specific
