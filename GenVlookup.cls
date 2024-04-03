'GenVlookup(シート名1, 検索値1セル, シート名2, 検索値2セル, 取得データセル, 数式組み込みセル)
'「シート名1の検索値1セルの列が、シート名2の検索値2セルの列と合致した時に、シート名2の取得データセルの
'列のデータをシート名1の数式組み込みセルの列に持ってくるようにVLOOKUPの数式を組み立てて入れる」
'※引数として入れるセル情報は、フィールド名直下の1件目のデータ
Sub GenVlookup(WS1 As String, Key1Cell As String, WS2 As String, Key2Cell As String, DataPickCell As String, FormulaCell As String)
    Dim DifCol_Relative As Variant
    Dim DifColNo As Long
    Dim KeyMain_R1C1 As Variant
    
    Dim RowNoBottom_Main As Long
    Dim RowNoBottom_Ref As Long
    Dim RowNo_KeyRef As Long
    Dim ColNo_KeyRef As Long
    Dim ColNo_DataPickCell As Long
    Dim FormulaToSet As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Worksheets(WS1).Activate
    DifCol_Relative = Range(Key1Cell).Column - Range(FormulaCell).Column
    KeyMain_R1C1 = "RC[" & DifCol_Relative & "]"
    
    RowNo_KeyRef = Range(Key2Cell).Row
    ColNo_KeyRef = Range(Key2Cell).Column
    ColNo_DataPickCell = Range(DataPickCell).Column
    
    RowNoBottom_Main = 汎用.Ck_EndxlDown0_F(WS1, Key1Cell)
    RowNoBottom_Ref = 汎用.Ck_EndxlDown0_F(WS2, Key2Cell)
    Worksheets(WS1).Activate
    DifColNo = Range(DataPickCell).Column - Range(Key2Cell).Column + 1
    FormulaToSet = FormulaCell & ":" & Split(Range(FormulaCell).Address, "$")(1) & RowNoBottom_Main
    Range(FormulaToSet) = _
    "=VLOOKUP(" & KeyMain_R1C1 & "," & WS2 & "!R" & RowNo_KeyRef & "C" & ColNo_KeyRef & _
    ":R" & RowNoBottom_Ref & "C" & DifColNo & "," & DifColNo & ",False)"
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
