'' --------------------------------------------------------------------------------
'' Sample for sheet BeforeDoubleClick refrence
'' --------------------------------------------------------------------------------

If Target.Row = 1 And Target.Value <> "" Then
SoryActiveRow
Exit Sub
End If

Dim text As String
Dim col As Integer

tTraget = Target.Value
CTarget = Target.Column

Call FilterByActiveCell(tTraget, CTarget)

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
    tbl.Range.AutoFilter Field:=iCol, Criteria1:=tVals

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