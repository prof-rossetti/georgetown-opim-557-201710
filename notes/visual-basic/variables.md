# VBA Language Overview

## Variables

### Defining Variables

Visual Basic has traditionally been a "statically-typed" language, which means it requires the developer to indicate as part of a variable's definition which [type of data](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/data-types) the variable will hold.

The most common way to define a variable is to use the `Dim` keyword, followed by the variable name, followed by the `as` keyword, followed by the datatype. For example:

```vb
Dim MyNumber as Integer
Dim MyText as String
Dim MyDecimal as Double
Dim MyBool as Boolean
```

After you study Excel Objects and ActiveX Controls, you can dynamically store them in variables by using the `Set` keyword instead of the `Dim` keyword:

```vb
Set MySheet = Application.ActiveSheet ' note: in addition to defining the variable, this also assigns it a value
```

### Assigning Values to Variables

Use an equality operator (`=`) to assign some value to a given variable. For example:

```vb
' For pre-defined variables:
MyNumber = 25
MyText = "Hello World"
MyDecimal = 3.14
```

### Referencing Variables

After variables are defined and assigned, any subsequent references to the variable name will yield the variable's value:

```vb
"All I have to say is: " & MyText ` --> "All I have to say is: Hello World"
```

```vb
MyNumber + MyDecimal ` --> 28.14
```

```vb
MySheet.Name ' --> "Sheet1"
```
