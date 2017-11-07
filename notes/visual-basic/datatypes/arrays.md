# VBA Language Overview

## Datatypes

### Arrays

Reference:

  + [Declaring Arrays](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/declaring-arrays)
  + [Using Arrays](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/using-arrays)

> Array: "A set of sequentially indexed elements having the same intrinsic data type. Each element of an array has a unique identifying index number. Changes made to one element of an array don't affect the other elements." - [glossary of VBA terms](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/vbe-glossary)

```vb
' declare an array, including its size and the type of data it will contain, if possible ...

Dim Teams(1 To 5) As String

' assign values to each position in the array ...

Teams(1) = "New York Yankees"
Teams(2) = "New York Mets"
Teams(3) = "Boston Red Sox"
Teams(4) = "New Haven Ravens"
Teams(5) = "Washington Nationals"

' access a given item in the array by referencing its position, or "index" ...

MsgBox( Teams(4) ) ' --> a message box displaying "New Haven Ravens"
```

See also: [looping](/notes/visual-basic/loops.md) through each item in an array.
