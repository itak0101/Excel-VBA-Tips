'---------------------------------------------------------------------------
' 選択範囲を画像ファイルとして出力する
'---------------------------------------------------------------------------
Sub OutputRange()
    
    ' 例外処理
    On Error GoTo Catch
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = False
    
    ' 選択範囲を画像として一度貼り付ける
    Range("A1:D4").CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ActiveSheet.Paste
    
    ' TEMPフォルダに選択範囲を出力(HTML形式)
    With ActiveWorkbook.PublishObjects.Add(xlSourceSheet, _
        Environ("TEMP") + "\img_get.htm", ActiveSheet.Name, "", xlHtmlStatic, "img", "")
        .Publish (True)
        .AutoRepublish = False
    End With

    ' TEMPフォルダから画像をコピー
    FileCopy _
        Source:=Environ("TEMP") + "\img_get.files\img_image001.png", _
        Destination:=ActiveWorkbook.Path + "\" + ActiveSheet.Name + ".png"

    ' シートに貼り付けた画像を削除
    Selection.Delete
    
    ' 警告メッセージボックスを表示に設定
    Application.DisplayAlerts = True
    
    Exit Sub

' 例外処理
Catch:
    MsgBox "例外が発生しました"
    Application.DisplayAlerts = True

End Sub
'---------------------------------------------------------------------------