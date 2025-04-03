' 換算表　250403 추가파일처리 결과 누적 
' 複数のExcelファイルを順次処理するメイン関数
Sub ProcessMultipleXLSXFiles()
    Dim csvFilePath As String
    Dim continueProcessing As Boolean
    Dim isFirstFile As Boolean
    
    ' まずCSVファイルを選択
    Call SelectCSVFile
    
    If g_csvFilePath = "" Then
        MsgBox "CSVファイルが選択されていません。処理を終了します。", vbExclamation
        Exit Sub
    End If
    
    csvFilePath = g_csvFilePath
    
    ' 結果シートを初期化（初回実行時）
    InitializeResultSheet
    
    ' 複数のExcelファイル処理ループ
    isFirstFile = True
    continueProcessing = True
    
    Do While continueProcessing
        ' Excelファイルを選択
        Call SelectXLSXFile
        
        If g_xlsxFilePath = "" Then
            ' ファイル選択がキャンセルされた場合
            Exit Do
        End If
        
        ' ファイル処理を実行
        ProcessFilesAndAppendResults csvFilePath, g_xlsxFilePath, isFirstFile
        
        ' 次の繰り返しからは最初のファイルではない
        isFirstFile = False
        
        ' 続けて処理するか質問する
        continueProcessing = (MsgBox("他のExcelファイルも処理しますか？", _
                                     vbQuestion + vbYesNo, "続けて処理") = vbYes)
    Loop
    
    ' 処理完了後、結果シートをアクティブ化してメッセージを表示
    ThisWorkbook.Sheets("結果").Activate
    
    MsgBox "すべてのファイル処理が完了しました。" & vbCrLf & _
           "結果は「結果」シートに表示されています。" & vbCrLf & _
           "内容を確認した後、「SaveResultSheetAsCSV」関数を実行して" & vbCrLf & _
           "結果をCSVファイルとして保存できます。", vbInformation
End Sub


' 結果シートを初期化する関数
Private Sub InitializeResultSheet()
    Dim resultSheet As Worksheet
    
    ' 既存の結果シートがあるか確認
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("結果")
    On Error GoTo 0
    
    ' シートがなければ新規作成
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "結果"
        
        ' 情報ラベル
        resultSheet.Cells(1, 1).Value = "保存フォルダ:"
        resultSheet.Cells(2, 1).Value = "ファイル名:"
        resultSheet.Cells(3, 1).Value = "処理状態:"
        
        ' 初期情報設定
        resultSheet.Cells(1, 2).Value = ThisWorkbook.Path
        resultSheet.Cells(2, 2).Value = "結果_統合_" & Format(Date, "yymmdd") & ".csv"
        resultSheet.Cells(3, 2).Value = "初期化済み"
        
        ' ヘッダー行の設定
        resultSheet.Cells(6, 1).Value = "処理ファイル"
        resultSheet.Cells(6, 2).Value = "登録番号"
        resultSheet.Cells(6, 3).Value = "L列データ"
        resultSheet.Cells(6, 4).Value = "M列データ"
        resultSheet.Cells(6, 5).Value = "マッチング状態"
        
        ' ヘッダー行の書式設定
        resultSheet.Range("A6:E6").Font.Bold = True
        resultSheet.Range("A6:E6").Interior.Color = RGB(220, 230, 241)
        
        ' 列幅自動調整
        resultSheet.Columns("A:E").AutoFit
    End If
End Sub

' ファイルを処理して結果を既存シートに追加する関数
Public Sub ProcessFilesAndAppendResults(csvFilePath As String, xlsxFilePath As String, Optional isFirstFile As Boolean = True)
    ' 変数宣言
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    Dim tempSheet As Worksheet  ' CSVデータ用の一時シート
    Dim filteredSheet As Worksheet  ' フィルタリングされたデータ用シート
    Dim resultSheet As Worksheet  ' 結果表示用シート
    
    ' パターン値の抽出
    Call ExtractPatternFromFilename(xlsxFilePath)
    
    ' ファイル名の抽出
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim xlsxFileName As String
    xlsxFileName = fso.GetFileName(xlsxFilePath)

    ' 進捗状況の表示
    Application.StatusBar = "ファイル処理中: " & xlsxFileName
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual  ' 計算を手動に設定して高速化
    
    ' 処理開始時間
    Dim startTime As Double
    startTime = Timer
    
    ' ここは既存のコードとほぼ同じで、CSVファイル処理と検索キーの作成を行います
    ' ...
    
    '====================================================================
    ' 3段階：結果を結果シートに追加
    '====================================================================
    
    ' 既存の結果シートを取得
    Set resultSheet = ThisWorkbook.Sheets("結果")
    
    ' 現在の結果シートの最終行を検索
    Dim lastResultRow As Long
    lastResultRow = resultSheet.Cells(resultSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 最初のデータ行インデックス設定
    If lastResultRow < 6 Then lastResultRow = 6
    
    ' 現在のExcelファイルに関する情報を更新
    If isFirstFile Then
        ' 最初のファイルの場合は処理状態を更新
        resultSheet.Cells(3, 2).Value = "処理中"
    End If
    
    ' 結果シートにデータ追加
    Dim rowsAdded As Long
    rowsAdded = 0
    
    ' 結果データの表示
    For i = 2 To lastRow
        ' L列とM列のデータを取得
        lValue = Trim(CStr(xlsWs.Cells(i, 12).Value))
        mValue = Trim(CStr(xlsWs.Cells(i, 13).Value))
        
        ' 結果シートの行インデックス計算
        Dim rowIdx As Long
        rowIdx = lastResultRow + rowsAdded + 1  ' 最終行の後に追加
        
        ' マッチング状態に応じて出力
        If matchStatus(i) Then
            ' マッチングした行 - 登録番号あり
            resultSheet.Cells(rowIdx, 1).Value = xlsxFileName  ' 処理ファイル名
            resultSheet.Cells(rowIdx, 2).Value = matchedRegNums(i)  ' 登録番号
            resultSheet.Cells(rowIdx, 5).Value = "マッチング"  ' 状態
        Else
            ' マッチングしなかった行 - 登録番号には初期値を設定
            resultSheet.Cells(rowIdx, 1).Value = xlsxFileName  ' 処理ファイル名
            resultSheet.Cells(rowIdx, 2).Value = "未マッチング-" & i  ' マッチングしていない識別子
            resultSheet.Cells(rowIdx, 5).Value = "未マッチング"  ' 状態
            
            ' マッチングしていない行の色指定
            resultSheet.Range("A" & rowIdx & ":E" & rowIdx).Interior.Color = RGB(255, 200, 200)
        End If
        
        ' L列とM列のデータ追加
        resultSheet.Cells(rowIdx, 3).Value = lValue
        resultSheet.Cells(rowIdx, 4).Value = mValue
        
        rowsAdded = rowsAdded + 1
    Next i
    
    ' 処理状態の更新
    Dim totalProcessed As Long
    totalProcessed = rowsAdded
    If Not isFirstFile Then
        ' 以前の処理データがある場合、合算
        Dim previousCount As String
        previousCount = resultSheet.Cells(3, 2).Value
        
        If InStr(previousCount, "行") > 0 Then
            ' "xxxx行 処理済み" 形式から数字を抽出
            totalProcessed = CLng(Split(Split(previousCount, "行")(0), " ")(0)) + rowsAdded
        End If
    End If
    
    resultSheet.Cells(3, 2).Value = totalProcessed & "行 処理済み"
    
    ' 結果シートの書式指定
    resultSheet.Columns("A:E").AutoFit
    resultSheet.Range("A6:E" & (lastResultRow + rowsAdded)).Borders.LineStyle = xlContinuous
    
    ' フィルタ設定 (まだない場合)
    If Not resultSheet.AutoFilterMode Then
        resultSheet.Range("A6:E6").AutoFilter
    End If
    
    ' 処理完了のメッセージ表示
    MsgBox """" & xlsxFileName & """ ファイルの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "処理時間: " & Format(processingTime, "0.00") & " 秒", vbInformation
End Sub

' 結果シートをCSVとして保存する関数
Public Sub SaveResultSheetAsCSV()
    ' "結果" シートを対象とする
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("結果")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "「結果」シートが見つかりません。処理を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' 省略...
    
    ' ADOStreamの設定
    With adoStream
        .Type = 2 ' テキストモード
        .Charset = "UTF-8"
        .Open
        
        ' ヘッダー行を直接書き込み
        .WriteText "処理ファイル,登録番号,L列データ,M列データ,マッチング状態", 1 ' 1=adWriteLine
        
        ' データ行の記録...
        ' 省略...
    End With
    
    ' 保存完了メッセージ
    MsgBox "結果ファイルが正常に保存されました:" & vbCrLf & resultFilePath, vbInformation
End Sub

' リソースを完全に解放する関数
Private Sub ThoroughCleanup(ByRef xlsApp As Object, ByRef xlsWb As Object, ByRef xlsWs As Object)
    On Error Resume Next
    
    ' ワークブックが設定されている場合は閉じる
    If Not xlsWb Is Nothing Then
        xlsWb.Close SaveChanges:=False
    End If
    
    ' Excelアプリケーションが設定されている場合は終了
    If Not xlsApp Is Nothing Then
        xlsApp.Quit
    End If
    
    ' オブジェクト参照の解放
    Set xlsWs = Nothing
    Set xlsWb = Nothing
    Set xlsApp = Nothing
    
    ' ガベージコレクションを強制実行
    GetObject("", "Excel.Application").Parent.Wait 1000  ' 小さい遅延
    CollectGarbage
End Sub

' ガベージコレクションを強制実行
Private Sub CollectGarbage()
    Dim i As Long
    For i = 1 To 2
        VBA.Collection  ' 空の参照を作成
    Next i
End Sub

