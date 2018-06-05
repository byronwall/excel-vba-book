# SO item 111
I'm working on my first VBA function. The goal is to have a function that accepts two integers as inputs, and outputs an array containing all the integers in between the two inputs (end-points included).

Example: If I input 5 and 9, the output should be an array of 5, 6, 7, 8, 9\.

VBA doesn't seem to have any of the objects or functions I'm used to in other languages. Python has a range() function, but most other languages I know about have list-like types which can be appended to. How does this work in VBA?

I'm not looking to create an Excel range, but rather an array which contains a range between two values.

----

**Applies only in the context of Excel** (not sure based on question and tags)

If you want a one-liner that doesn't deal with arrays, you can take advantage of the `ROW` function and `Application.Evaluate` within Excel.

**Code**

```
Function RangeArr(int_start As Integer, int_end As Integer) As Variant

    RangeArr = Application.Transpose( _
                    Application.Evaluate("=ROW(" & int_start & ":" & int_end & ")"))

End Function

```

**Results**

![array results](https://i.stack.imgur.com/F5jPB.png)
