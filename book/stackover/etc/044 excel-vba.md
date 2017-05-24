# SO item 044
I have data of varying character length in column `AJ` (anywhere from 3 to 5, or 0 characters). I use the following piece of code to display the character length in column AK.

```
Range("AK3:AK15000").Formula = "=LEN(AJ3)"

```

I want create a new column of data in column AL such that if the character length in AK is 5, it will return what was in column AJ. If the character length in AK is 4, I need it to return what was in column AJ, but with a zero (0) in front. If the character length is 3, I need it to return what was in column AJ, but with two zeros (00) in front. If the character length is zero (0), then return nothing. The following code will not work and returns an error.

```
If Range("AK3:AK15000").Value = 5 Then
    Range("AL3:AL15000").Formula = "=AJ3"
ElseIf Range("AK3:AK15000").Value = 3 Then
    Range("AL3:AL15000").Formula = "CONCATENATE(0,0,AJ3)"
ElseIf Range("AK3:AK15000").Value = 4 Then
    Range("AL3:AL15000").Formula = "CONCATENATE(0,AJ3)"
End If

```

Can someone please help me?

----

You can also make good use of `REPT` here and have it work for any length string less than the pad length. It will error out if the string is longer than the pad length, but it sounds like that won't happen based on your description.

Here is an example with data in `AJ4`.

```
=IF(AJ4<>"",REPT(0,5-LEN(AJ4))&AJ4,"")

```

Change `5` to whatever length you want it padded to.
