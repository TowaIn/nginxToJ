Sub AddGFTNextToNDB()
    ' 変数の宣言
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetCells As Collection
    Dim i As Long

    ' --- 初期設定 ---
    ' アクティブなシートを対象に設定
    Set ws = ActiveSheet
    ' 処理対象となるセルを格納するコレクションを初期化
    Set targetCells = New Collection

    ' 画面の更新を一時的に停止（処理の高速化のため）
    Application.ScreenUpdating = False

    ' --- 「NDB」セルの検索と収集 ---
    ' シート内でデータが存在する範囲（UsedRange）を検索対象に設定
    Set searchRange = ws.UsedRange
    
    ' 範囲内の各セルをチェック
    For Each foundCell In searchRange
        ' セルの値が「NDB」と完全に一致する場合
        If foundCell.Value = "NDB" Then
            ' 処理対象としてコレクションに追加
            targetCells.Add foundCell
        End If
    Next foundCell

    ' --- 「GFT」の挿入処理 ---
    ' 収集したセルを後ろから処理（セルの挿入による位置ずれを防ぐため）
    If targetCells.Count > 0 Then
        For i = targetCells.Count To 1 Step -1
            Set foundCell = targetCells(i)
            
            ' 見つかったセルの右隣にセルを1つ挿入（データは右へシフト）
            foundCell.Offset(0, 1).Insert Shift:=xlToRight
            
            ' 新しくできた右隣のセルに「GFT」と入力
            foundCell.Offset(0, 1).Value = "GFT"
        Next i
    End If

    ' --- 終了処理 ---
    ' 画面の更新を再開
    Application.ScreenUpdating = True

    ' 処理完了のメッセージを表示
    MsgBox "処理が完了しました。 " & targetCells.Count & "件の「NDB」の隣に「GFT」を挿入しました。"

End Sub
