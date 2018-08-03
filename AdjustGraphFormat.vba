'-------------------------------------------------------------------------------
' グラフの体裁を整える
'-------------------------------------------------------------------------------
Sub AdjustGraphFormat()

    '-- 宣言と初期化 ---------------------------------------------
    Dim nVar As Integer
    Dim dVar As Double
    Dim sVar As String
    Application.DisplayAlerts = False ' 警告メッセージボックスを非表示に設定

    '-- 全シートのループ -----------------------------------------
    For i = 1 To Worksheets.Count
    
        ' シートを一つ取得
        Set TargetSheet = Worksheets(i)
        TargetSheet.Activate
    
        ' シート名やA1セルに記載されている内容で処理分岐
        If TargetSheet.Name <> "GraphSheet" Then
        
            'シート内の全グラフのループ
            For j = 1 To TargetSheet.ChartObjects.Count
                
                ' グラフを一つ取得
                Set chartObj = TargetSheet.ChartObjects(j)
                Set Chart = chartObj.Chart
                
                ' グラフ名の設定
                chartObj.Name = "Graph1"
                
                ' データ範囲の設定 (F4セルに"A1:B10"といった記載がある想定)
                If TargetSheet.Range("F4").Value <> "" Then
                    sVar = TargetSheet.Range("F4").Value
                    Chart.SetSourceData Source:=Range(sVar)
                End If
                
                ' タイトルの設定
                If TargetSheet.Range("F5").Value <> "" Then
                    Chart.HasTitle = True
                    sVar = TargetSheet.Range("F5").Value
                    Chart.ChartTitle.Text = sVar
                End If
                
                ' x軸ラベルの設定
                If TargetSheet.Range("F6").Value <> "" Then
                    Chart.Axes(xlCategory).HasTitle = True
                    sVar = TargetSheet.Range("F6").Value
                    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = sVar
                End If
                
                ' y軸ラベルの設定
                If TargetSheet.Range("M7").Value <> "" Then
                    Chart.Axes(xlValue).HasTitle = True
                    sVar = TargetSheet.Range("L7").Value
                    Chart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = sVar
                End If
                
                ' x軸最小値の設定
                If TargetSheet.Range("F8").Value <> "" Then
                    dVar = TargetSheet.Range("F8").Value
                    Chart.Axes(xlCategory).MinimumScale = dVar
                End If
                
                ' x軸最大値の設定
                If TargetSheet.Range("F9").Value <> "" Then
                    dVar = TargetSheet.Range("F9").Value
                    Chart.Axes(xlCategory).MaximumScale = dVar
                End If
                
                ' y軸最小値の設定
                If TargetSheet.Range("F10").Value <> "" Then
                    dVar = TargetSheet.Range("F10").Value
                    Chart.Axes(xlValue).MinimumScale = dVar
                End If
                
                ' y軸最大値の設定
                If TargetSheet.Range("F11").Value <> "" Then
                    dVar = TargetSheet.Range("F11").Value
                    Chart.Axes(xlValue).MaximumScale = dVar
                End If
                
                ' 縦サイズの設定
                If TargetSheet.Range("F12").Value <> "" Then
                    dVar = TargetSheet.Range("F12").Value
                    chartObj.Height = dVar
                End If
                
                ' 横サイズの設定
                If TargetSheet.Range("F13").Value <> "" Then
                    dVar = TargetSheet.Range("F13").Value
                    chartObj.Width = dVar
                End If

            Next j
    
        End If
        
    Next i ' -----------------------------------------

    ' 最後に1シート目を開き、次回起動時に1シート目が表示されるようにする
    Worksheets(1).Activate

End Sub

'-------------------------------------------------------------------------------