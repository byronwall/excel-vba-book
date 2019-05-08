## declaring the return type (Function only)

For a Function, the only extra step is to declare the return type of the Function. This is done after the normal parameters, with an extra `as Type` where `Type` is the actual type that you want to return. Note that this type must be compatible with all possible Types that you could return. Sometimes this means that you need to return a Variant in order to have all possible return Types available to you. There are times where this makes sense (and a large part of the Excel object model does this), but note that using Variant will make it hard to use Intellisense to figure out what your VBA is capable of doing.

TODO: is this a Variant by default?

TODO: give some examples of Function returns (or link to examples of them)
