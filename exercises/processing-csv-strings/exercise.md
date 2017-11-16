# "Processing CSV Strings" Exercise

## Prerequisites

  + [Arrays](/notes/visual-basic/datatypes/arrays.md)
  + [Splitting Strings](/notes/visual-basic/datatypes/strings.md#string-splitting)

## Learning Objectives

  + Practice parsing strings which exist in comma-separated values (CSV) format.
  + Practice looping through an array of items.
  + Practice using a loop to programmatically write cell values.

## Challenge

**Write VBA code that will process the following Comma-separated Values (CSV) string into a corresponding spreadsheet of cells.**

Desired input (`MyStr`):

```vb
Dim MyStr As String

MyStr = "city,name,league" & vbNewLine & _
        "New York,Mets,Major" & vbNewLine & _
        "New York,Yankees,Major" & vbNewLine & _
        "Boston,Red Sox,Major" & vbNewLine & _
        "Washington,Nationals,Major" & vbNewLine & _
        "New Haven,Ravens,Minor"

MsgBox(MyStr)

' write some VBA code here!
```

Desired output (spreadsheet of cells):

city | name | league
--- | --- | ---
New York | Mets | Major
New York | Yankees | Major
Boston | Red Sox | Major
Washington | Nationals | Major
New Haven | Ravens | Minor

## Walkthrough

TBA!
