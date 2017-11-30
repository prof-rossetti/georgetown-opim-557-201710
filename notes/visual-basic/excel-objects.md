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
Worksheets("Sheet1") ' specify the sheet name

Worksheets("Sheet1").Range("A1:C5") ' specify the address of a range of cells
```

More commonly you can specify object references relative to the active workbook or worksheet:

```vb
Worksheets("Sheet1").Range("A1:C5")

Range("A1:C5")
```

### The `Range` Object

The `Range` object represents one or more cells.

#### Helpful Range Properties

##### Reading Cell Values

Read the value of a cell:

```vb
Dim MyVar As String
MyVar = Range("A1").Value
MsgBox("The value in cell A1 is: " & MyVar)
```

Alternative approach to referencing cell properties:

```vb
Dim MyCell As Range
Set MyCell = Range("A1")
MsgBox("The value in cell " & MyCell.Address & " is: " & MyCell.Value)
```

See also [Datatypes of Numeric Cell Values](/notes/visual-basic/datatypes.md#datatypes-of-numeric-cell-values).

##### Writing Cell Values

Write a value to a cell:

```vb
Range("A1").Value = "fun times"
```

##### Clearing Cell Contents

Clear the contents of some range:

```vb
Range("A1:C5").ClearContents
```

##### Cells in a Range

Access all cells in a given range:

```vb
Range("A1:C5").Cells.Count ' --> 15
```

After studying [loops](/notes/visual-basic/loops.md#for-each--next-loops), you can use one to iterate through all cells in a given range.

### The `Worksheet` Object

The `Worksheet` object references a corresponding worksheet. Access a specific sheet by its name (e.g. "Sheet1") or its position in the workbook (e.g. 1).

```vb
Dim MySheet As Worksheet
Set MySheet = Worksheets("Sheet1") ' or Worksheets(1) if this is the first sheet

MySheet.Name ' --> "Sheet1"
MySheet.Index ' --> 1
MySheet.Activate ' switch user view to this sheet
```

Like the `Range` object, the `Worksheet` object also has a `Cells` property, which can be used to manipulate the sheet's cell values.

```vb
MySheet.Cells.ClearContents ' remove values of all cells in this sheet
```

Pass a row number and a column number to reference a specific cell:

```vb
MySheet.Cells(1,3).Value = "good stuff" ' where 1 is the row number and 3 is the column number (a.k.a. cell "C1")
```

#### Creating and Deleting Sheets

[Create](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/sheets-add-method-excel) a new sheet:

```vb
ActiveWorkbook.Worksheets.Add
```

[Delete](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/worksheet-delete-method-excel) a given sheet:

```vb
Worksheets("Sheet3").Delete
```

#### Looping Through Sheets

After you have studied looping and arrays, you can apply the concepts to loop through a collection of worksheets (`ActiveWorkbook.Worksheets`):

```vb
Dim MySheet As Worksheet

For Each MySheet In ActiveWorkbook.Worksheets ' adapted from source: https://www.extendoffice.com/documents/excel/3057-excel-delete-all-sheets-except-current.html
    MsgBox (MySheet.Name)
Next
```
