'---------------------------------------------------------------------------
' 指定範囲を画像ファイルとして出力する (全シート)
'---------------------------------------------------------------------------
Sub OutputRangeAllSheet()
    
    ' ファイル操作系クラスを定義
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' 出力フォルダの作成 (既に存在している場合は一度削除してから新規作成)
    sOutputFolderPath = ActiveWorkbook.Path + "\OutputRange"
    If FSO.FolderExists(sOutputFolderPath) Then
        FSO.DeleteFolder (sOutputFolderPath)
    End If
    FSO.CreateFolder (sOutputFolderPath)
    
    ' 全シートループ
    For i = 1 To Worksheets.Count
    
        ' 対象シート情報の取得
        Set TargetSheet = Worksheets(i) ' 処理対象シートを1つ取得
        sSheetName = TargetSheet.Name   ' 処理対象シート名を取得する
        TargetSheet.Activate            ' 処理対象シートを選択する
        
        ' 対象に処理実施
        Call OutputRange(sSheetName, "A1:D4")
        
    Next i
    
    ' 画像出力時の中間ファイルを削除
    If FSO.FolderExists(sOutputFolderPath + "\image.files") Then
        FSO.DeleteFolder (sOutputFolderPath + "\image.files")
    End If
    If FSO.FileExists(sOutputFolderPath + "\image.htm") Then
        FSO.DeleteFile (sOutputFolderPath + "\image.htm")
    End If
    
    ' 最後に1シート目を開いてから終了することで、次回起動時に1シート目が表示されるようにする
    Worksheets(1).Activate
    
End Sub


'---------------------------------------------------------------------------
'  指定範囲を画像ファイルとして出力する (対象シート)
'---------------------------------------------------------------------------
Sub OutputRange(ByVal sSheetName As String, ByVal sRange As String)
    
    ' 例外処理
    On Error GoTo Catch
    
    ' 処理対象シートをアクティブにする
    Worksheets(sSheetName).Activate
    
    ' 処理対象シートのA1セルを選択する
    Worksheets(sSheetName).Range("A1").Select
    
    ' 出力先フォルダが存在しなければエラー終了
    sOutputFolderPath = ActiveWorkbook.Path + "\OutputRange"
    If Dir(sOutputFolderPath, vbDirectory) = "" Then
        Err.Raise Number:=999, Description:="画像出力フォルダが存在しません。事前に作成してください。" & vbNewLine & sOutputFolderPath
    End If
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = False
        
    ' 選択範囲を画像としてExcel上に貼り付ける
    Range(sRange).CopyPicture Appearance:=xlScreen, Format:=xlPicture
    ActiveSheet.Paste
    
    ' Excel上に貼り付けた画像をファイル出力(1次出力)
    With ActiveWorkbook.PublishObjects _
        .Add(xlSourceSheet, sOutputFolderPath + "\image.htm", ActiveSheet.Name, "", xlHtmlStatic, "AAA", "")
        .Publish (True)
        .AutoRepublish = False
    End With

    ' 1次出力先フォルダから最終出力先にファイルをコピー(シート内に図形が含まれる場合は001を002などに変更する必要あり)
    FileCopy _
        Source:=sOutputFolderPath + "\image.files\AAA_image001.png", _
        Destination:=sOutputFolderPath + "\" + ActiveSheet.Name + ".png"

    ' Excel上に貼り付けた画像を削除
    Selection.Delete
    
    ' 警告メッセージボックスを表示に設定
    Application.DisplayAlerts = True
    
    Exit Sub

' 例外処理
Catch:
    MsgBox (Err.Description)
    Application.DisplayAlerts = True

End Sub


'---------------------------------------------------------------------------



