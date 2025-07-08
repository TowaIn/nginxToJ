Sub DuplicateColumnAndReplaceAll()
    ' 変数の宣言
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim j As Long
    Dim newCol As Range
    Dim cell As Range

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
            
            ' (3) 新しくできた列（j+1列目）を変数に設定
            Set newCol = ws.Columns(j + 1)
            
            ' (4) 新しい列の全セルをチェックして "NDB" を "GFT" に置換
            '   (UsedRange との重複範囲のみを対象とし、処理を高速化)
            For Each cell In Intersect(newCol, ws.UsedRange)
                ' セル内に "NDB" という文字列が含まれているかチェック
                If InStr(1, cell.Value, "NDB", vbTextCompare) > 0 Then
                    ' "NDB" を "GFT" に置換する
                    cell.Value = Replace(cell.Value, "NDB", "GFT", 1, -1, vbTextCompare)
                End If
            Next cell
            
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
