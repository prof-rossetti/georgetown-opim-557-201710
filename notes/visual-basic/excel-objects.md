# VBA Language Overview

## MS Excel Objects

Documentation:

  + [`Application`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/application-object-excel)
  + [`Workbook`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/workbook-object-excel)
  + [`Worksheet`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheet-object-excel)
  + [`Range`](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-object-excel)

### Collections

Use collections to reference a specific excel object by its identifying characteristic.

You can specify absolute object references:

```vb
Application.Workbooks("my-book.xlsm") ' specify the workbook's file name

Application.Workbooks("my-book.xlsm").Worksheets("Sheet1") ' specify the sheet name

Application.Workbooks("my-book.xlsm").Worksheets("Sheet1").Range("A1:C5") ' specify the address of a range of cells
```

More commonly you can specify object references relative to the active workbook or worksheet:

```vb
Worksheets("Sheet1").Range("A1:C5")

Range("A1:C5")
```

### The `Range` Object

The `Range` Object represents one or more cells.

#### Example Code

Clear the contents of some range:

```vb
Range("A1:C5").ClearContents
```

Read the value of a cell:

```vb
Dim MyVar As String
MyVar = Range("A1").Value
MsgBox("The value in cell A1 is: " & MyVar)
```

Write a value to a cell:

```vb
Range("A1").Value = "fun times"
```

Alternative approach to referencing cell properties:

```vb
Set MyCell = Range("A1") ' important to use Set instead of Dim here
MsgBox("The value in cell " & MyCell.Address & " is: " & MyCell.Value)
```
