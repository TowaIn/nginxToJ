Sub DuplicateColumnAndRename()
    ' 変数の宣言
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim j As Long

    ' --- 初期設定 ---
    ' アクティブなシートを対象に設定
    Set ws = ActiveSheet

    ' 画面の更新を一時的に停止（処理の高速化のため）
    Application.ScreenUpdating = False

    ' --- 列の検索と処理 ---
    ' 4行目の最終列を取得
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column

    ' 列を右側から左側へ向かってチェック (列挿入によるループずれを防ぐため)
    For j = lastCol To 1 Step -1
        
        ' 4行目のセルの値が "NDB" の場合
        If ws.Cells(4, j).Value = "NDB" Then
            
            ' (1) "NDB" がある列全体をコピー
            ws.Columns(j).Copy
            
            ' (2) 右隣にコピーした列を挿入
            ws.Columns(j + 1).Insert Shift:=xlToRight
            
            ' (3) 新しくできた列（j+1列目）の4行目の値を "GFT" に変更
            ws.Cells(4, j + 1).Value = "GFT"
            
        End If
    Next j

    ' --- 終了処理 ---
    ' コピーモードを解除 (点滅する破線を消す)
    Application.CutCopyMode = False
    
    ' 画面の更新を再開
    Application.ScreenUpdating = True

    ' 処理完了のメッセージを表示
    MsgBox "処理が完了しました。"

End Sub
