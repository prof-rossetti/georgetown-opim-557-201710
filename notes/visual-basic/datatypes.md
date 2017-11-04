# VBA Language Overview

## Datatypes

Reference: [documentation](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/data-types).

Popular datatypes include:

  + `Integer` - a positive whole number
  + [`String`](datatypes/strings.md) - text
  + `Double` - a decimal number
  + `Boolean` - true or false
  + `Date` - a calendar date
  + `Array` - an ordered collection of items

You can also think about each [Excel Object](/notes/visual-basic/excel-objects.md) as belonging to its own datatype.

### Checking a Variable's Type

Visual Basic supports a number of built-in functions to detect the datatype of any variable. These functions are especially helpful when validating user inputs.

[The `TypeName()` function and `TypeOf ... Is` statement](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/early-late-binding/determining-object-type) returns a String value to indicate the datatype.

```vb
TypeName(1) ' --> "Integer"
TypeName("Hello") ' --> "String"
TypeName(True) ' --> "Boolean"
TypeName(3.14) ' --> "Double"
TypeName(#10/31/2017#) ' --> "Date"
TypeName(Range("A1:A7")) ' --> "Range"

' perform a string comparison
If TypeName("Hello") = "String" Then
  MsgBox("'Hello' is a string datatype")
End If
```

[The `VarType()` function](https://support.office.com/en-us/article/VarType-Function-1e08636c-1892-40c2-aff3-2b894389e82d) "returns an Integer indicating the subtype of a variable". See the function's reference document for a table mapping the resulting integers to corresponding datatypes.

```vb
VarType(1) ' --> 2
VarType("Hello") ' --> 8
VarType(True) ' --> 11
VarType(3.14) ' --> 5
VarType(#10/31/2017#) ' --> 7
VarType(Range("A1:A7")) ' --> 8204

' perform an integer comparison
If VarType("Hello") = 8 Then
  MsgBox("'Hello' is a string datatype")
End If
```

[The `IsNumeric()` function](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/isnumeric-function) evaluates whether or not a variable "can be evaluated as number".

```vb
IsNumeric(1) ' --> True
IsNumeric(3.14) ' --> True
IsNumeric(True) ' --> True
IsNumeric("1") ' --> True
IsNumeric("3.14") ' --> True
IsNumeric("Hello") ' --> False
```
