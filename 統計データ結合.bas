Sub 統計データ結合()
    ' 変数宣言
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, fileCount As Long
    Dim headerRow As Long
    Dim dataArr As Variant
    
    ' フォルダ選択ダイアログを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "データファイルのあるフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されていません。プログラムを終了します。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' 新しいワークブックを作成
    Set targetWb = Workbooks.Add
    Set targetWs = targetWb.Sheets(1)
    targetWs.Name = "結合データ"
    
    ' 最初の行にヘッダーを設定
    targetWs.Cells(1, 1).Value = "ファイル名"
    
    ' フォルダ内のxlsxファイルを処理
    fileName = Dir(folderPath & "\*.xlsx")
    fileCount = 0
    
    ' 結果ワークシートの初期行
    lastRow = 1
    
    ' ファイルが存在する間、処理を繰り返す
    Do While fileName <> ""
        ' 現在のファイル名をスキップ
        If fileName <> targetWb.Name Then
            fileCount = fileCount + 1
            
            ' ファイルを開く
            Set wb = Workbooks.Open(folderPath & "\" & fileName, ReadOnly:=True)
            Set ws = wb.Sheets(1) ' 最初のシートを使用
            
            ' ファイル内のデータ範囲を取得
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            lastFileRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' データを配列に読み込む
            dataArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastFileRow, lastCol)).Value
            
            ' 最初のファイルの場合、ヘッダーを取得
            If fileCount = 1 Then
                headerRow = 1
                ' ヘッダーをコピー (B列から)
                For j = 1 To lastCol
                    targetWs.Cells(1, j + 1).Value = dataArr(1, j)
                Next j
            Else
                headerRow = 2 ' ヘッダーをスキップ
            End If
            
            ' データをコピー
            For i = headerRow To lastFileRow
                ' ファイル名を A列に設定
                targetWs.Cells(lastRow + i - headerRow + 1, 1).Value = fileName
                
                ' 元のデータを B列以降にコピー
                For j = 1 To lastCol
                    targetWs.Cells(lastRow + i - headerRow + 1, j + 1).Value = dataArr(i, j)
                Next j
            Next i
            
            ' 次のデータの位置を更新
            lastRow = lastRow + lastFileRow - headerRow + 1
            
            ' ファイルを閉じる
            wb.Close SaveChanges:=False
        End If
        
        ' 次のファイルを取得
        fileName = Dir()
    Loop
    
    ' 結果を整形
    targetWs.UsedRange.Columns.AutoFit
    
    ' 処理完了メッセージ
    MsgBox fileCount & "個のファイルが正常に結合されました。", vbInformation
    
End Sub
