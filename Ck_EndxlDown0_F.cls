Function Ck_EndxlDown0_F(WS As String, CellName As String)
'*** When the the cell in the row is emply while reading records from top to bottom,this procedure regards it as the end of the records.
Sheets(WS).Activate
    If Range(CellName).Offset(1, 0) <> "" Then
        Ck_EndxlDown0_F = Range(CellName).End(xlDown).Row  '=Cells(Row_, Col_).End(xlDown).row
    Else
        Ck_EndxlDown0_F = Range(CellName).Row
    End If
End Function
