'-------------------------------------------------------------------------------
' 全シートのグラフを画像出力する
'-------------------------------------------------------------------------------
Sub ExportGraph()

    ' 宣言と初期化
    Dim i As Integer
    Dim SheetName As String
    Dim SumSheetName As String
    Dim iss As Boolean
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = False

    ' 画像出力フォルダが存在しなければ作成する
    SaveDir = ThisWorkbook.Path & "\\GraphOut_" & Format(Now, "yyyymmdd_hhmmss")
    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If
    
    '全シートのループ --------------------------------------
    For i = 1 To Worksheets.Count
    
        ' シートを1つ取得
        Set TargetSheet = Worksheets(i)
        TargetSheet.Activate
        SheetName = TargetSheet.Name
    
        ' シート名やA1セルに記載されている内容で処理分岐
        If SheetName <> "" Then
        'If TargetSheet.Range("A1").Value = "AAAAA" Then
        
            'シート内の全グラフのループ
            For j = 1 To TargetSheet.ChartObjects.Count
                
                ' グラフを1つ取得
                Set chartObj = TargetSheet.ChartObjects(j)
                Set Chart = chartObj.Chart
                
                ' ファイル出力
                chartObj.Select
                ActiveChart.Export (SaveDir & "\\" & SheetName & "-" & j & ".png")

            Next j
    
        End If
        
    Next i ' -----------------------------------------------
    
    '1番目のシートにカーソルを合わせる
    Worksheets(1).Activate
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = True
    
End Sub
'-------------------------------------------------------------------------------