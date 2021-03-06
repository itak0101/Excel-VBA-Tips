'---------------------------------------------------------------------------
' 全シートでA1セルにカーソルを合わせる
'---------------------------------------------------------------------------
Sub SelectA1()
    ' 宣言と初期化
    Dim TargetSheet As Worksheet
    
    ' 全シートループ
    For i = 1 To Worksheets.Count
        Set TargetSheet = Worksheets(i) ' 処理対象シートを1つ取得
        TargetSheet.Activate            ' 処理対象シートを選択する
        TargetSheet.Range("A1").Select  ' 処理対象シートのA1セルを選択する
    Next i
    
    ' 最後に1シート目を表示して終わり
    Worksheets(1).Activate
End Sub
'---------------------------------------------------------------------------