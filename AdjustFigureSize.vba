'-------------------------------------------------------------------------------
' 全シートの図をサイズ変更する
'-------------------------------------------------------------------------------
Sub AdjustFigureSize()

    ' 宣言と初期化 ------------------------------
    Dim i As Integer
    Dim SheetName As String
    Dim SumSheetName As String
    Dim TargetSheet As Worksheet
    Dim TargetShape As Shape
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = False
    '--------------------------------------------
    
    '全シートのループ --------------------------------
    For i = 1 To Worksheets.Count
    
        ' シートを一つ取得
        Set TargetSheet = Worksheets(i)
        TargetSheet.Activate
        SheetName = TargetSheet.Name
    
        ' シート名が「Agenda」でなければ処理実行
        If (SheetName <> "Agenda") Then
            For j = 1 To TargetSheet.Shapes.Count
            
                ' 図表オブジェクトを取得
                Set TargetShape = TargetSheet.Shapes(j)
                
                ' 図表オブジェクトが図であった場合
                If InStr(TargetShape.Name, "Picture") Then
                    TargetShape.LockAspectRatio = True    '縦横比を固定
                    TargetShape.Width = 600               '幅を指定(ポイントで指定)
                    ' TargetShape.Width =Application.CentimetersToPoints(12.35) '幅を指定(cmで指定)
                End If
            Next j
        End If
    Next i ' -----------------------------------------
    
    
    '1番目のシートにカーソルを合わせる
    Worksheets(1).Activate
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = True
    
End Sub

