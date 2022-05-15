## Excel Notes

The following VB run a simple query of a table and gathers the id's needed.

Nothing secret about it but might be useful to people.

The function below can also beused in a formular:

e.g. `=gather_ids_r(Sheet2!$B$4:$B$16,Sheet2!$C$4:$C$16,C7,Sheet2!$D$4:$D$16,D7)`

```visualbasic
Sub QueryMacro1()
'
' Macro1 Macro
'

'
    Dim ids As String
    Dim offsetcell As Range
    Dim resultcell As Range
    
    Dim R As Range
    
    ' Get the current selection of the current active cell
    Set CurrentRegion = Selection
    For Each currentcell In Selection
        ' Debug.Print ("Selection: " & currentcell.Value & " " & currentcell.Address)
        ' Get the value of the cell one to right of the current active cell
        Set offsetcell = currentcell.Offset(0, 1)
        ' Reference the Data worksheet
        Set W = Worksheets("Sheet2")
        ' Pull the range from th worksheet above
        Set R = W.Range("B4", W.Range("B4").End(xlDown))
        ' Query for the ids
        ids = gather_ids( _
            R, _
            1, currentcell, _
            2, offsetcell _
        )
        ' place the ids found in the cell to the left of the selected column
        Set resultcell = currentcell.Offset(0, -1)
        resultcell.Value = ids
    Next
End Sub

Function gather_ids_r( _
        ByVal query_table As Range, _
        ByVal test_col_1 As Range, ByVal query_test_1 As Range, _
        ByVal test_col_2 As Range, ByVal query_test_2 As Range _
    ) As String
    Dim ids As String
    Dim separator As String
    Dim val As String
    Dim test_1 As String
    Dim test_2 As String
    
    separator = ","
    Index = 0
    
    For Each idcell In query_table
        val = idcell.Value
        For i = 1 To 10
            Set can = idcell.Offset(0, i)
            If Not Application.Intersect(can, test_col_1) Is Nothing Then
                test_1 = can.Value
            End If
            If Not Application.Intersect(can, test_col_2) Is Nothing Then
                test_2 = can.Value
            End If
        Next
        
        'test_2 = idcell.Offset(0, test_col_2).Value
        If test_1 = query_test_1.Value And test_2 = query_test_2.Value Then
            ids = ids & val & separator
        End If
    Next
    If Len(ids) <> 0 Then
        ids = Left(ids, Len(ids) - Len(separator))
    End If
    gather_ids_r = ids
End Function

Function gather_ids( _
        ByVal query_table As Range, _
        ByVal test_col_1 As Integer, ByVal query_test_1 As Range, _
        ByVal test_col_2 As Integer, ByVal query_test_2 As Range _
    ) As String
    Dim ids As String
    Dim separator As String
    Dim val As String
    Dim test_1 As String
    Dim test_2 As String
    
    separator = ","
    Index = 0
    
    For Each idcell In query_table
        val = idcell.Value
        test_1 = idcell.Offset(0, test_col_1).Value
        test_2 = idcell.Offset(0, test_col_2).Value
        If test_1 = query_test_1.Value And test_2 = query_test_2.Value Then
            ids = ids & val & separator
        End If
    Next
    If Len(ids) <> 0 Then
        ids = Left(ids, Len(ids) - Len(separator))
    End If
    gather_ids = ids
End Function


```
