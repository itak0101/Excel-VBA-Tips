'----------------------------------------------------------------------------------------
' 現在開いているシートでソルバーを実行する
'----------------------------------------------------------------------------------------
Sub Fitting(ByVal sRange_Para1 As String, ByVal sRange_Para2 As String, ByVal sRange_Error As String)
' Function エラーが表示される場合は、
' VBAマクロ画面→「ツール」→「参照設定」→「Solver」にチェックを入れてください。

    ' Solver設定データの初期化
    SolverReset
                
    ' 制約条件を設定する
    ' CellRef 制約を課すセル
    ' Relation 等号指定 1:「指定セル≦境界値」、2:「指定セル=境界値」、3:「指定セル≧境界値」
    ' FormulaText 境界値
    SolverAdd CellRef:=sRange_Para1, Relation:=3, FormulaText:="0"
    SolverAdd CellRef:=sRange_Para1, Relation:=1, FormulaText:="100"
    SolverAdd CellRef:=sRange_Para2, Relation:=3, FormulaText:="0"
    SolverAdd CellRef:=sRange_Para2, Relation:=1, FormulaText:="100"
                  
    '上記制約条件の元でフィッティングを実行
    ' SetCell パラメータ値の適正度を表現する値 (例: 残差二乗和)
    ' MaxMinVal 1は最大化、2は最小化、3は特定の値と一致
    ' ByChange 最適化したいパラメータ (例: 線形フィッティングの場合、傾きaと切片b)
    ' Engine フィッティング方法 1:GRG非線形、2:シンプレックスLP、3:エヴォリューショナリー
    SolverOk SetCell:=sRange_Error, MaxMinVal:=2, ByChange:=(sRange_Para1 & "," & sRange_Para2), Engine:=1
        
    '解析後の確認画面を非表示設定にする
    SolverSolve UserFinish:=True

End Sub

'----------------------------------------------------------------------------------------
' グラフ1についてフィッティングを実行する
'----------------------------------------------------------------------------------------
Sub Fitting_Graph1()

    ' G16セルに傾き、G17セルに切片、G18セルに残差二乗和が記載されているイメージ
    Call Fitting("G16", "G17", "G18")

End Sub

'----------------------------------------------------------------------------------------
' グラフ2についてフィッティングを実行する
'----------------------------------------------------------------------------------------
Sub Fitting_Graph2()

    ' K16セルに中央値、K17セルに対数標準偏差、K18セルに残差二乗和が記載されているイメージ
    Call Fitting("K16", "K17", "K18")

End Sub

'----------------------------------------------------------------------------------------
