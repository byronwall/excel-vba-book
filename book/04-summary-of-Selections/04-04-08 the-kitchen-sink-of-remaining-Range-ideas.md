### the kitchen sink of remaining `Range` ideas

- Pull the `Range` reference from some other object
- Name a cell and use that name indirectly -- `Names("CellName")`
- Ask the user to select the `Range` to use
- Use a function to get a reference -- `Application.Index`
- Search for the cell based on its function or value -- `Find()`
- Process a formula to determine the `Range` it depends on

TODO: look into the Trace functions to see what they return

#### Objects that will return a Range

One of the greatest consistencies throughout VBA and the Object Model is how various objects will return a new object or reference to a useful property. At times, this can save you a large chunk of time trying to recreate that access from scrath. The key then is knowing when these properties exist and how to use them.

Below is a rough summary of objects that will give you access to a Range.

- TODO: create this list
- TODO: consider making this a cheat sheet or similar since it covers most of the sections in this chapter

In addition to objects that will return a Range, there are also objects which will not return a Range but should. These include:

- TODO: create the rest of this list
- Chart Series info related to the Name, Values, and XValues. You are required to work through the `=SERIES` formula instead

#### Using `Names().RefersToRange`

There are two ways to work with named ranges. One of them is quite simple: `Range("SomeNamedRange")`. This works well in a couple of cases:

- You know the exact name you want to use or can prompt the user for it
- You are using the `Range` call on an object that has proper scope.

For the latter point, the default named ranges have `Workbook` scope and the `Range` call works across the board. This becomes more of an issue when you are using the same name across multiple Worksheets with a Worksheet level scope. You can still access the named ranged, but now your call to `Range`, needs to be `Worksheet.Range` from the correctly scoped Worksheet.

The former ppint about needing to know the name is more often the problem. Sometimes you want to help someone use a named range, but you simply do not know what they are named. One trivial example is creating an addin that outputs all of the named ranges in the Workbook. You cannot iterate them through `Range` because you want to know what they are!

When you are in a position where you want to use the named ranges but do not know or want to use the actual names, you can go directly through the `Names` object. There are two ways to do this:

- Iterate the `Names` with no knowledge of them
- Use an index, i.e. the `Name` and call into `Names(index)`

Once you have access to a valid `Name`, you can then access the `RefersToRange` which will return a Range that can be used. There are few instances where this is ever going to be better if you already have the name. The one exception to this is if you are wanting to change some of the metadata associated with the Name. This mainly includes the comment on the name since there is not much else. another option is that you can copy the named Range as a new range with a slightly different name. I have done this before to process all of the named ranges into some new named Range based on a formula which included the previous one. This can be a critical step to improving the performance of array formulas that previously pointed to entire columns. The problem is that create the dynamically named ranges is an absolute pain without VBA.

TODO: add an example of the dynamic name creation

Once you are comfortable accessing named ranges, you may find that it is helpful to create them from time to time from VBA. This can be a helpful way of storing a complicated Range that your VBA created without having to select the cells and hope you can type the name correctly.

#### Using `Application.InputBox(, Type:=8)`

One very useful technique for obtaining a Range is to ask the user for one. This is one of the fastest ways to level up your VBA game because it provides the user control while also making your VBA look pretty slick with the Range picker. The other upside here is that the InputBox Range picker generally works better than the RedEdit version on a form. The odd thing here is two-fold:

- You have to know that InputBOx exists on the Application alone. IF you use the other version, then you cannot supply the Type
- YOu have to know that Type:=8 allows for a Range selection

ONce oyu have two those things down (because you read this book!) then you are able to ask the user to pick a Range with ease. The other very nice thing about the InputBox approach is that you can supply a default address (not Range) and it will automatically be selected at the start. I have used this approach to get effect in bUTL to allow the VBA to process the Selection (by default) or to allow the user to select something different. This is a very clean solution to sneivle defaults while also allowing the user to do something different once they read your initial prompt. It is also dead simple to upgrade your current `Set rng = Range()` to `Set rng = Application.InputBox("Select a cell", Type:=8)` instead. For utility type code, the difference in immense in terms of not having to hard code or guess Ranges. Or you can still guess them but provide the user a chance to change the guess.

TODO: move that Function here form bUTL GetOrSelect...

#### Using `Application.Index`

The `=INDEX` formula is the most potent formula in Excel. Its counterpart in the VBA world is also powerful but less impressive compared to real programming. Having said that, the `Index` function works exactly as expected in VBA and is a very nice tool to have if you are comfortable using INDEX in a normal spreadsheet. The real power of Index is that you can use it to replace a lot of the common code where you iterate through a Range until you find given value. One potential upside of Index is that you can upgrade an Excel only methodology over to VBA with minial change to formulas. Once you have the work converted over, you can then set about addin the details that VBA alone can provide.

TODO: does this work any different than Cells? is it really that useful?

#### Using `Range.Find()`

I seldom use `Range.Find()`, but it can be a powerful addition when you know what you want to search for. My problem with .Find is that it is incredibly rare that I have some free text I am searching for and want to find using VBA. Generally speaking, Find becomes useful when you are processing a somewhat arbitrary Worksheet which may contain certain data you want. In my experience, I am far more likely to use an AutoFilter or something other than Find. Part of the problem for me is that I have never had a problem using some other method than Find. I also generally find myself somewhat confused by the parameters and the general executon of Find. Typically, you will need to create a While loop to search for the next found items.

I also have the (probably unfair) view that Find is a crutch to not being able to use other methods to Find a given Range. I generally prefer to iterate through cells and check values. My mind is built around building a Range and processing it rather than attempting to find a Range and then process it. Your mileage may vary.

TODO: add an example of using Find correctly

#### Pulling a Range from a Formula with string processing

One of the next level things to do with VBA is to start processing your Formulas to drive your VBA. There are a couple of places where this might be useful:

- You are dealing with a Chart Series Formula which must be parsed
- You want to Trace the precedent cells but don't want to deal with TracePrecedents
- You want to modify some part of the formula (e.g. take `A1` and surround it with an `ABS(A1)`)
- You want to make all of the cells in a specific formula a specific color (like a permanent version of hitting `F2`)

Whatever your motivation, it's good to remember that the formulas in a spreadsheet are generally the most important information aside from the actual data. IN some spreadsheets, the formulas are the only important part. If you want to extract and use this information, then it is helpful to be able to parse the formulas and identify the Ranges.

There are a couple of approaches to parsing Ranges from formulas, depending on what you need to do and what you start with:

- Your formulas contain only A1 style references without sheet names
- Your formulas may contain a sheet name too
- You want to extract non-range formula information

For the first two, you can build relatively simple parsers which can extract the Range information which good accuracy. The key here is to understand exactly what your formulas look like. The worst case is having to build a full out formulas parser which is a non-trivial exercise. Hadnling all possible Excel syntaxs is a mess.

If you can settle for something less, then you have a couple of approaches at hand:

- Use a Regular Expression keyed in to Range options
- Use your knowledge of the possible formulas to extract the relevant parts with string functions

TODO: add an example of some Regex which work here... expanding complexity

TODO: add an example of parsing out with Split and Left or something
