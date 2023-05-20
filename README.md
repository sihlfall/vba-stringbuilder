# VBA / VB6 StringBuilder

## Motivation
Building strings by repeated concatenation is slow. Whenever two strings are concatenated, space for a new string will be
allocated and the contents of the existing string will be copied into the newly allocated space.
Thus, in the worst case, namely when assembling strings of equal length, one ends up with an algorithm that is proportional to the _square_ of the number of strings.

The goal of this project is to provide something for VBA that is similar to the  _StringBuilder_ class in .NET. The design goals are:

 * _Performance:_ We focus on the core functionality (appending strings) and leave out any convenience functions that would increase the work to be done.
 * _Portability:_ We stick to a VB-only solution, i.e. we do not refer to any Windows DLLs.

Our StringBuilder comes in two flavours: The module `StaticStringBuilder` defines a user-defined type `StaticStringBuilder.Ty`.
Its intended use case is a StringBuilder
being needed locally. Variables of type `StaticStringBuilder.Ty` should not be assigned to each other
and values should only passed by reference, since assignment or passing
by value will induce copying. For use cases in which a StringBuilder  needs to be accessed from various parts of the application,
we provide a `StringBuilder` class, objects of which must be
created dynamically with `New`.

`StaticStringBuilder` is a few percent faster than `StringBuilder`.

Both flavours use an exponentially growing buffer in order to achieve near-linear performance. In the worst case, for assembling _n_ strings of equal length, time complexity is O(_n_ log _n_), i.e. near linear.

## Usage

You need _either_ `src/StaticStringBuilder.bas` _or_ `src/StringBuilder.cls`.

### StaticStringBuilder

Copy or import `StaticStringBuilder.bas` into your application.

Usage:
```
Dim sb As StaticStringBuilder.Ty
StaticStringBuilder.AppendStr sb, "First"
StaticStringBuilder.AppendStr sb, "Second"
StaticStringBuilder.AppendStr sb, "Third"

Dim s As String
s = StaticStringBuilder.GetStr(sb)
' Now s will contain the string "FirstSecondThird".
```

### StringBuilder

Copy or import `StringBuilder.cls` into your application.

Usage:
```
Dim sb As StringBuilder
Set sb = New StringBuilder
sb.AppendStr "First"
sb.AppendStr "Second"
sb.AppendStr "Third"

Dim s As String
s = sb.Str
' Now s will contain the string "FirstSecondThird".
```

## Performance comparison

We compare

* _naive_ repeated string concatenation using VB's `&` operator,
* a _StringBuilder_ class by VolteFace and Dragokas from [here](https://www.vbforums.com/showthread.php?847365-VB6-StringBuilder-Fast-string-concatenation) and [here](https://github.com/sancarn/stdVBA-Inspiration/tree/master/Better%20StringBuilder), which relies heavily on native Windows DLL functions for accessing and managing memory,
* our _StaticStringBuilder_,
* our _StringBuilder_ class.

The following table shows the time it takes to append _n_ single character strings to an initial empty string, for different values of _n_ (all values in ms). Naturally, the time figures are to be understood as relative values—actual absolute times will depend on the machine and the environment. Our measurements were made on a relatively recent machine (as of spring 2023) with an executable generated by VB 6:

|_n_                 |10000|20000|50000|100000|200000|
|:---                |  --:|  --:|  --:|   --:|   --:|
|naive               |1.79 | 6.50|41.92|163.25|622.15|   
|VolteFaceDragokas   |0.55 | 1.05| 2.67|  5.29| 10.54|
|StaticStringBuilder |0.38 | 0.72| 1.84|  3.68|  7.51|
|StringBuilder class |0.53 | 1.03| 2.63|  5.32| 10.69|

From the results, one can observe that both StringBuilders show runtimes that are comparable to a library using native DLLs.

## Varia

### Cute features of the StringBuilder class

For convenience, the `StringBuilder` class supports the following features, inspiration for which was drawn from the
[stdVBA project](https://github.com/sancarn/stdVBA):

* The property `Str` is the default property, which means the string held by the StringBuilder can be obtained or set
  by simple assignment. Thus, if you have a StringBuilder and a string,
  ```
  Dim sb As StringBuilder, s As String
  ```
  you can write
  ```
  s = sb  ' equivalent to s = sb.Str
  sb = s  ' equivalent to sb.Str = s
  ```
  For this to work in VBA, the class has to be _imported_ (rather than copied) into the project, since the feature relies on function attributes, and while VBA _does_ support function attributes, there is no way of _defining_ them in the VBA editor.

* `AppendStr` is defined to be the class's _Evaluate_ method, which enables calling the function by simply enclosing the argument in
  square brackets _if_ the StringBuilder object is late-bound, i.e. if the variable is declared as being of type `Object`:
  ```
  Dim sb As Object, s As String
  Set sb = New StringBuilder
  sb.[First]
  sb.[Second]
  sb.[Third]
  s = sb
  ' Now s will contain the string "FirstSecondThird".
  ```
  Like the previous feature, this relies on function attributes, so it will only work if the class is imported. Note, however, that late binding costs some percents of performance.

### Design ideas

We avoid frequent memory allocations and copying by pre-allocating a buffer of some predefined minimum capacity (the default is
16 characters). Whenever the capacity of the buffer is exhausted, we increase the buffer size by 50% by allocating a new, larger buffer, and copying the contents from the old buffer over to the new buffer subsequently freeing the old buffer.

In VBA, Strings are mutable (their contents can be modified by left-hand side `Mid`); thus we can use a string variable for holding the buffer. Allocating a certain capacity is realized by creating a string with the corresponding number of placeholder characters. As the StringBuilder fills up, these placeholder characters are replaced by the actual characters of the string being built.

There is one small caveat, though: When we increase the capacity of the buffer, the two buffers must exist simultaneously—the old one, from which the data is copied, as well as the new one, to which the data is copied. Since in VBA, each assignment of string variables induces a copy, we will end up with one unnecessary copy of an entire buffer, either when storing the old buffer in the temporary variable or when reading the populated new buffer from the temporary variable. In order to avoid this, a StringBuilder contains an array of two buffer string variables, which are used alternately. Apart from the process of a capacity change, at any point in time, only one of the buffer variables is active, while the other one is null.   





