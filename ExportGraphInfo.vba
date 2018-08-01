'-------------------------------------------------------------------------------
' Excel内に配置されている全グラフの情報をファイル出力する
'-------------------------------------------------------------------------------
Sub ExportGraphInfo()

    '出力ファイルオープン
    outFile = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & "_GraphInfo.csv"
    Open outFile For Output As #1
  
    '出力(ヘッダ行)
    Print #1, ActiveWorkbook.Path & "\" & ActiveWorkbook.Name
    Print #1, "シート名,グラフ名,グラフタイトル,X軸ラベル,Y軸ラベル,系列名,系列タイトル,系列データ";
     
    '[Loop] 全シート
    For i = 1 To Worksheets.Count
        Set TargetSheet = Worksheets(i)
        
        '[Loop] シート内の全オブジェクト
        For j = 1 To TargetSheet.Shapes.Count
            Set TargetShape = TargetSheet.Shapes(j)
            
            '[分岐A] オブジェクトがグラフの場合
            If InStr(TargetShape.Name, "Chart") Or InStr(TargetShape.Name, "Graph") Then
                Set TargetChart = TargetSheet.Shapes(j).Chart
                
                '[Loop] グラフ内の全系列
                On Error Resume Next 'ループ内で例外が発生したら、無視して次の項目に進む
                For k = 1 To TargetChart.SeriesCollection.Count
                    Set TargetCollection = TargetChart.SeriesCollection(k)
                    
                    '出力(改行のみ)
                    Print #1, ""
                    
                    '出力(数式以外の部分、改行なし)
                    Print #1, _
                        TargetSheet.Name _
                        & ",""" & TargetShape.Name & """" _
                        & ",""" & TargetChart.ChartTitle.Text & """" _
                        & ",""" & TargetChart.Axes(xlValue).AxisTitle.Text & """" _
                        & ",""" & TargetChart.Axes(xlCategory).AxisTitle.Text & """" _
                        & ",""" & TargetChart.Name & """" _
                        & ",""" & TargetCollection.Name & """";
                    
                    '出力(数式部分、改行なし)
                    Print #1, ",""" & Replace(TargetCollection.Formula, " = ", "") & """";
                        
                Next k '[Loop] グラフ内の全系列
                
            End If '[分岐A] オブジェクトがグラフの場合
            
        Next j '[Loop] シート内の全オブジェクト
        
    Next i '[Loop] 全シート
    
    '出力ファイルクローズ
    Close #1
    
    '終了通知
    MsgBox "ファイル出力しました" & vbCrLf & outFile

End Sub
'-------------------------------------------------------------------------------