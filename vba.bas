'' --------------------------------------------------------------------------------
'' Sample for sheet BeforeDoubleClick refrence
'' --------------------------------------------------------------------------------

If Target.Row = 1 And Target.Value <> "" Then
SoryActiveRow
GoTo EndHere
End If

Dim text As String
Dim col As Integer

tTraget = Target.Value
CTarget = Target.Column

Call FilterByActiveCell(tTraget, CTarget)

EndHere:
Cancel = True
Target.Select

'' --------------------------------------------------------------------------------

Sub FilterByActiveCell(tVal, iCol)

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim tName As String
  
    tName = ActiveCell.ListObject.Name
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(tName)
    
    If tbl.AutoFilter.FilterMode = True Then
        tbl.AutoFilter.ShowAllData
        Exit Sub
    End If

    With tbl.Range
        .AutoFilter
        .AutoFilter Field:=iCol, Criteria1:=tVal
    End With

End Sub

'' --------------------------------------------------------------------------------


Sub SoryActiveRow()

Dim ws As Worksheet
    Dim tbl As ListObject

    Dim tName As String
  
    tName = ActiveCell.ListObject.Name
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(tName)
    Set rng = Range(tbl & "[" & ActiveCell & "]")
    
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
    
End Sub

'' --------------------------------------------------------------------------------
Function getActiveTable(tbl As ListObject)

  Dim ws As Worksheet
  Dim tName As String
  Dim tCount As Integer

  Set ws = ActiveSheet

  '' check for only 1 Table in a sheet
  '' ------------------------------------------
  For Each tbl In ws.ListObjects
  tName = tbl.Name
  tCount = tCount + 1
      If tCount > 1 Then
          MsgBox ("More then 1 Table in Sheet - End Sub")
          Exit Function
      End If
  Next tbl
  '' ------------------------------------------

  Set getActiveTable = ws.ListObjects(tName)
  
End Function

Function getLastRowID()

  Dim lastRow As Integer
  Dim lastID As Integer
  Dim tbl As ListObject

  Set tbl = getActiveTable(tbl)


  lastRow = tbl.Range.Rows.Count
  lastID = Cells(lastRow, 1)
  
  Let getLastRowID = lastID

End Function

Sub addNewLog()

Dim tbl As ListObject
Dim newrow As ListRow
Dim rowId As Integer

Set tbl = getActiveTable(tbl)
rowId = getLastRowID() + 1
Set newrow = tbl.ListRows.Add
With newrow
    .Range(1) = rowId
    .Range(8) = Now()
End With


End Sub
