## Controlling calculations

When you are creating macro workflows, there are a number of tools at your disposal to control calculations flow. Before describing those tools, it's worth stepping back and discussing why you might want to control the calculation flow. There are a couple of common reason:

- Performance. Your code will run faster if you control the calculation process. This mainly involves disabling automatic calculation at key points.
- Accuracy. For some types of calculatiojns, you need to tightly control the calculation flow for accuracy. This is often the case if you are building a spreadsheet that does some form of recursion or self reference.
- Usability. There are some situations where you are interacting with calculations and need to prevent the normal behavior. The most common is when you add Workbook events like `Change`.
- Profiling. If you are building a code profiler (i.e. a tool that tracks execution time of your code) you must control calculations in order to get the tracking right.

We'll get back to the applciations, but it's also worht hittin gthe high points on how you can control the calculation. THe main knobs:

- Disable application wide
- Disable for a Worksheet
- Manually calculate a Range, Worksheet, or Application

THe types of changes you will make are fiarly tightly couple to the applciations above. In general, for performances nad usability reasons, you will be disable calculations. For accuracy or profiling applcaitions, you will manually walking the calculation through.
