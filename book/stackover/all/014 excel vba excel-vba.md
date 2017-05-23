# SO item 014
I have a range

```
Set rng = Range("B5:H20")

```

I want to create a subrange that contains all cells except for those on the first row of `rng`. What is a good way to do this?

```
Set subRng = 'Range("B6:H20")

```

----

Another versatile approach to this is the `Offset` and `Intersect` pattern. It has the advantage over `Resize` of working the same regardless of how you do the shift (could move 1 column or row w/o rethinking the `Resize` part).

It also technically works for discontinuous ranges although a use case for that is rare at best.

```
Set subrng = Intersect(rng.Offset(1), rng)

```
