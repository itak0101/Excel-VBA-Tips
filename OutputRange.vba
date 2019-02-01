' 選択範囲をPNG画像として保存する
Sub OutputRange()
    
    ' 例外処理
    On Error GoTo Catch
    
    ' 警告メッセージボックスを非表示に設定
    Application.DisplayAlerts = False
    
    ' 選択範囲を画像として一度貼り付ける
    Range("A1:D4").CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ActiveSheet.Paste
    
    ' 選択範囲を1次出力
    With ActiveWorkbook.PublishObjects _
        .Add(xlSourceSheet, ActiveWorkbook.Path + "\images\image.htm", ActiveSheet.Name, "", xlHtmlStatic, "AAA", "")
        .Publish (True)
        .AutoRepublish = False
    End With

    ' 1次出力フォルダからコピー
    FileCopy _
        Source:=ActiveWorkbook.Path + "\images\image.files\AAA_image001.png", _
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
