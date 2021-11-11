Sub FilterByActiveCell(tVal, iCol)

    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim tName As String
  
    tName = ActiveCell.ListObject.Name
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(tName)
    
    If tbl.AutoFilter.FilterMode = True Then
        tbl.AutoFilter.ShowAllData
    End If
    tbl.Range.AutoFilter Field:=iCol, Criteria1:=tVal

End Sub

'' --------------------------------------------------------------------------------
'' Sample for sheet BeforeDoubleClick refrence
'' --------------------------------------------------------------------------------

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

Dim text As String
Dim col As Integer

tTraget = Target.Value
CTarget = Target.Column

Call FilterByActiveCell(tTraget, CTarget)

Cancel = True
Target.Select

End Sub
'' --------------------------------------------------------------------------------

Sub SoryByID()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim tName As String
    Dim ColName As String
    Dim SortCol As String
    
    tName = ActiveCell.ListObject.Name
    ColName = "ID" '' Insert table column title
    
    Set ws = ActiveSheet
    Set tbl = ws.ListObjects(tName)
    SortCol = tName & "[" & ColName & "]"
    
    Set rng = Range(SortCol)
                   
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
    
End Sub