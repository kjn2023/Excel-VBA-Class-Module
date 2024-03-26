'***************************************************************************
'※　FldNameCell_はあらかじめソートしておく
'クラスモジュールの「汎用」に"DelDup(FldNameCell_ As String)"を登録し、
'標準モジュールかフォームモジュールのプロシージャから以下のように実行
'Option Explicit
'Dim 汎用 as New 汎用
'Sub DelDup1()
'  Call 汎用.DelDup("B2")  '引数は重複時にレコードを削除する列のフィールド名のセル
'End Sub
'***************************************************************************
Sub DelDup(FldNameCell_ As String)
    Dim RowTop_ As Integer          '範囲内最上行
    Dim RowBottom_ As Integer       '範囲内最下行
    Dim Rng As Range
    Dim Rec1 As Variant: Rec1 = ""  '先に読むレコード
    Dim Rec2 As Variant: Rec2 = ""  '後に読むレコード
    Dim i As Integer                'for～Nextのカウンター
    Dim Col As Integer              'A列から数えた列番号
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Col = Range(FldNameCell_).Column
    
    Set Rng = Range(FldNameCell_).CurrentRegion.Item(Col)
    RowTop_ = Rng.Row
    RowBottom_ = Rng.End(xlDown).Row
    
    For i = RowTop_ To RowBottom_
        Rec1 = Cells(i, Col)
        If Rec1 = "" And Rec2 = "" Then
            Exit For
        End If
        
        If Rec1 = Rec2 Then
            Cells(i, Col).EntireRow.Delete
            i = i - 1
        End If
        Rec2 = Rec1
    Next

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
