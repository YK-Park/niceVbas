'CopyandRename 250201-01
' フォルダから Excel ファイル名を取得するサブプロシージャ
Sub GetExcelFileNames()
    Dim folderPath As String
    Dim fileName As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim fd As Office.FileDialog
    
    ' フォルダ選択ダイアログを表示
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Excel ファイルが格納されているフォルダを選択してください"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "処理を中止しました。", vbInformation
            Exit Sub
        End If
    End With
    
    ' 結果を書き込むシートを作成または取得
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ファイル一覧")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "ファイル一覧"
    End If
    On Error GoTo 0
    
    ' ヘッダーを設定
    ws.Cells(1, 1).Value = "ファイル名"
    ws.Cells(1, 2).Value = "パス"
    
    ' 既存のデータをクリア
    ws.Range("A2:B" & ws.Rows.Count).Clear
    
    ' 行カウンター初期化
    rowNum = 2
    
    ' フォルダ内の Excel ファイルを検索
    fileName = Dir(folderPath & "\*.xlsx")
    Do While fileName <> ""
        ' ファイル名を A 列に書き込み
        ws.Cells(rowNum, 1).Value = Left(fileName, InStrRev(fileName, ".") - 1)
        ' フルパスを B 列に書き込み
        ws.Cells(rowNum, 2).Value = folderPath & "\" & fileName
        
        rowNum = rowNum + 1
        fileName = Dir()
    Loop
    
    ' xlsm ファイルも検索
    fileName = Dir(folderPath & "\*.xlsm")
    Do While fileName <> ""
        ' ファイル名を A 列に書き込み
        ws.Cells(rowNum, 1).Value = Left(fileName, InStrRev(fileName, ".") - 1)
        ' フルパスを B 列に書き込み
        ws.Cells(rowNum, 2).Value = folderPath & "\" & fileName
        
        rowNum = rowNum + 1
        fileName = Dir()
    Loop
    
    ' 列幅を自動調整
    ws.Columns("A:B").AutoFit
    
    ' 結果を表示
    If rowNum > 2 Then
        MsgBox "ファイル一覧を取得しました。" & vbNewLine & _
               "取得件数: " & rowNum - 2 & "件", vbInformation
        
        ' シートをアクティブにする
        ws.Activate
    Else
        MsgBox "指定されたフォルダに Excel ファイルが見つかりませんでした。", vbExclamation
    End If
End Sub

' ファイルをコピーして名前を変更するサブプロシージャ
Sub CopyFiles()
    Dim wsNames As Worksheet
    Dim sourceFile As String
    Dim outputPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim prefix As String
    Dim suffix As String
    Dim errorCount As Long
    Dim successCount As Long
    
    ' 変数の初期化
    errorCount = 0
    successCount = 0
    
    ' コピー元ファイルの選択
    sourceFile = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
                                          MultiSelect:=False, _
                                          Title:="コピー元となるファイルを選択してください")
    
    ' キャンセルされた場合は終了
    If sourceFile = "False" Then
        MsgBox "処理を中止しました。", vbInformation
        Exit Sub
    End If
    
    ' プレフィックスを入力
    prefix = Application.InputBox("ファイル名の前につける文字列を入力してください", "プレフィックス設定", "2024年度_研修レポート_")
    If prefix = "False" Then Exit Sub
    
    ' 名前リストのシートを設定
    Set wsNames = ThisWorkbook.Sheets("名簿")
    
    ' 出力フォルダのパスを設定
    outputPath = ThisWorkbook.Path & "\output\"
    
    ' 出力フォルダが存在しない場合は作成
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    ' 最終行を取得
    lastRow = wsNames.Cells(wsNames.Rows.Count, "A").End(xlUp).Row
    
    ' 進捗状況を表示
    Application.ScreenUpdating = False
    
    ' 各名前に対してファイルをコピー
    For i = 2 To lastRow
        On Error Resume Next
        
        Dim name As String
        Dim newFileName As String
        name = wsNames.Cells(i, 1).Value
        
        If name <> "" Then
            newFileName = prefix & name & suffix & ".xlsx"
            FileCopy sourceFile, outputPath & newFileName
            
            If Err.Number = 0 Then
                successCount = successCount + 1
                Debug.Print "コピー完了: " & newFileName
            Else
                errorCount = errorCount + 1
                Debug.Print "エラー: " & newFileName & " - " & Err.Description
            End If
        End If
        
        Err.Clear
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "ファイルコピーが完了しました。" & vbNewLine & _
           "成功: " & successCount & "件" & vbNewLine & _
           "エラー: " & errorCount & "件", _
           IIf(errorCount = 0, vbInformation, vbExclamation)
End Sub


Sub UpdateNames()
    Dim wsNames As Worksheet
    Dim outputPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim errorCount As Long
    Dim successCount As Long
    Dim unlockedRange As Range  ' 保護を解除する範囲
    
    ' 変数の初期化
    errorCount = 0
    successCount = 0
    
    ' 名前リストのシートを設定
    Set wsNames = ThisWorkbook.Sheets("名簿")
    
    ' 出力フォルダのパスを設定
    outputPath = ThisWorkbook.Path & "\output\"
    
    ' 最終行を取得
    lastRow = wsNames.Cells(wsNames.Rows.Count, "A").End(xlUp).Row
    
    ' 進捗状況を表示
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 各ファイルに対して処理を実行
    For i = 2 To lastRow
        On Error Resume Next
        
        Dim name As String
        Dim fileName As String
        name = wsNames.Cells(i, 1).Value
        
        If name <> "" Then
            fileName = Dir(outputPath & "*" & name & "*.xlsx")
            
            If fileName <> "" Then
                ' ファイルを開く
                Set wb = Workbooks.Open(outputPath & fileName)
                
                If Err.Number = 0 Then
                    ' Sheet1が存在するか確認
                    On Error Resume Next
                    Set ws = wb.Sheets("Sheet1")
                    
                    If Not ws Is Nothing Then
                        ' シートの保護を解除
                        If ws.ProtectContents Then
                            ws.Unprotect
                        End If
                        
                        ' すべてのセルをロック
                        ws.Cells.Locked = True
                        
                        ' 特定の範囲のロックを解除
                        ' 例：D4セルとE4:G4範囲のロックを解除
                        Set unlockedRange = ws.Range("D4,E4:G4")
                        unlockedRange.Locked = False
                        
                        ' オプション：特定の範囲の数式を表示しない設定
                        ws.Cells.FormulaHidden = False  ' デフォルトですべての数式を表示
                        
                        ' 名前を入力
                        ws.Range("D4").Value = name
                        
                        ' シートを保護（ユーザーが実行可能な操作を設定）
                        ws.Protect _
                            UserInterfaceOnly:=True, _
                            Contents:=True, _
                            DrawingObjects:=True, _
                            Scenarios:=True, _
                            AllowFormattingCells:=False, _
                            AllowFormattingColumns:=False, _
                            AllowFormattingRows:=False, _
                            AllowInsertingColumns:=False, _
                            AllowInsertingRows:=False, _
                            AllowInsertingHyperlinks:=False, _
                            AllowDeletingColumns:=False, _
                            AllowDeletingRows:=False, _
                            AllowSorting:=False, _
                            AllowFiltering:=False, _
                            AllowUsingPivotTables:=False
                        
                        ' ファイルを保存して閉じる
                        wb.Save
                        wb.Close
                        
                        successCount = successCount + 1
                        Debug.Print "更新完了: " & fileName
                    Else
                        errorCount = errorCount + 1
                        Debug.Print "シートなし: " & fileName
                    End If
                Else
                    errorCount = errorCount + 1
                    Debug.Print "ファイルオープンエラー: " & fileName
                End If
            End If
        End If
        
        Err.Clear
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "名前の更新が完了しました。" & vbNewLine & _
           "成功: " & successCount & "件" & vbNewLine & _
           "エラー: " & errorCount & "件", _
           IIf(errorCount = 0, vbInformation, vbExclamation)
End Sub