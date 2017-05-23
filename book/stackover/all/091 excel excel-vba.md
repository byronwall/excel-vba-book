# SO item 091
I am trying to find a way to add status to my project automatically. For example, when cell 1 in Column A is at 50% then i would want to change cell 2 in the B column to "In Progress". Is this possible in Excel?

please let me know. thanks

----

Yes. A simple formula will do

```
=IF(A1=0.5, "In Progress", "")

```

This assumes that 50% is stored as 0.5\. You can also do `>=` if that makes more sense.
