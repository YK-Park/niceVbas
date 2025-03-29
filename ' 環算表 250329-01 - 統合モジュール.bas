' 環算表 250329-01 - 統合モジュール
Option Explicit

'====================================================================
' 1. グローバル変数と定数
'====================================================================

' グローバル変数
Public g_csvFilePath As String
Public g_xlsxFilePath As String
Public g_resultFilePath As String

' ファイル名から抽出したパターン値
Public g_fourDigits As String
Public g_oneDigit As String

' ワークシートセルの位置定義
Public Const CELL_STATUS = "B10"  ' 処理ステータス

'====================================================================
' 2. 初期化と基本機能
'====================================================================


' CSVファイル選択
Sub SelectCSVFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "CSVファイルを選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            g_csvFilePath = .SelectedItems(1)
        End If
    End With
End Sub

' XLSXファイル選択
Sub SelectXLSXFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "XLSXファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            g_xlsxFilePath = .SelectedItems(1)
        End If
    End With
End Sub

'====================================================================
' 3. ユーティリティ関数
'====================================================================

' ファイル名から4桁の数字-1桁の数字のパターンを抽出する関数
Public Function ExtractPatternFromFilename(filename As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String
    baseName = fso.GetFileName(filename)
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = "(\d{4})-(\d)"
        .Global = True
        .IgnoreCase = True
        
        Dim matches As Object
        Set matches = .Execute(baseName)
        
        If matches.Count > 0 Then
            ' マッチした場合は、4桁の数字と1桁の数字を取得
            g_fourDigits = matches(0).SubMatches(0)
            g_oneDigit = matches(0).SubMatches(1)
            ExtractPatternFromFilename = g_fourDigits & "-" & g_oneDigit
        Else
            g_fourDigits = ""
            g_oneDigit = ""
            ExtractPatternFromFilename = ""
        End If
    End With
End Function

' XLSXファイル名が条件を満たしているか確認する関数
Public Function IsValidXlsxFilename(filename As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String
    baseName = fso.GetFileName(filename)
    
    ' 条件1: "データ"または"処理"が含まれているか
    If InStr(1, baseName, "データ") > 0 Or InStr(1, baseName, "処理") > 0 Then
        IsValidXlsxFilename = True
        Exit Function
    End If
    
    ' 条件2: 4桁の数字-1桁の数字のパターン
    Dim pattern As String
    pattern = ExtractPatternFromFilename(filename)
    
    If pattern <> "" Then
        IsValidXlsxFilename = True
    Else
        IsValidXlsxFilename = False
    End If
End Function

' マッピングテーブルから実際の値を取得する関数
Public Function GetMappedValue(codeValue As String, codeCol As String, valueCol As String, startRow As Integer, endRow As Integer) As String
    ' コード値から実際の値に変換する
    ' codeValue: コード値
    ' codeCol: マッピングシートのコード列（例: "A"）
    ' valueCol: マッピングシートの実際の値列（例: "B"）
    ' startRow/endRow: 検索行範囲
    
    Dim mappingSheet As Worksheet
    Set mappingSheet = ThisWorkbook.Worksheets("参照")
    
    Dim i As Integer
    For i = startRow To endRow
        If CStr(mappingSheet.Range(codeCol & i).Value) = CStr(codeValue) Then
            GetMappedValue = CStr(mappingSheet.Range(valueCol & i).Value)
            Exit Function
        End If
    Next i
    
    ' マッチしない場合は元の値を返す
    GetMappedValue = codeValue
End Function

' 登録番号から各部分を抽出する関数
Public Function ExtractRegistrationParts(regNum As String) As Object
    Dim parts(0 To 7) As Variant
    
    parts(0) = ""
    parts(1) = ""
    parts(2) = ""
    parts(3) = ""
    parts(4) = False

    If Len(regNum) < 20 Then
        ' 長さが不足している場合はすぐに戻る
        Set ExtractRegistrationParts = parts
        Exit Function
    End If
    ' 固定位置から各部分を抽出
    ' 例: regNum = "ABC-D01232010101FRIA" の場合
    parts(0) = Mid(regNum, 6, 4)    ' "0123"
    parts(1) = Mid(regNum, 10, 2)   ' "01"
    parts(2) = Mid(regNum, 12, 7)   ' "0101FRI"
    parts(3) = Mid(regNum, 19, 1)   ' "A"
    parts(4) = True
    
    Set ExtractRegistrationParts = parts
End Function


'====================================================================
' 4. ファイル処理関数
'====================================================================

' CSVファイルをワークシートにインポートし、X列のデータを読む簡単な方法
Public Function ReadCSVToSheet(csvFilePath As String) As Worksheet

    On Error GoTo ErrorHandler

    ' CSVデータを保存するためのシートを準備
    Dim ws As Worksheet
    Dim wsExists As Boolean
        wsExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "CSV" Then
            wsExists = True
            Exit For
        End If
    Next ws


    ' 既存のシートがあれば削除
    If wsExists Then
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("CSVデータ").Delete
    Application.DisplayAlerts = True
    End If
    
    ' 新しいシートを作成
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "CSVデータ"

    ' CSVファイルを開く
    Dim lineText As String
    Dim lineItems() As String
    Dim fileNum As Integer
    Dim rowCount As Long
    
    fileNum = FreeFile
    Open csvFilePath For Input As #fileNum
    
    ' ヘッダー行を設定
    ws.Cells(1, 1).Value = "登録番号"
    rowCount = 2  ' データは2行目から
    
    ' CSVファイルを行ごとに処理
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        lineItems = Split(lineText, ",")
        
        ' X列（24列目）のデータがあるか確認
        If UBound(lineItems) >= 23 Then  ' 0ベースなので24列目は添字23
            ' 引用符を削除
            Dim cellValue As String
            cellValue = lineItems(23)  ' X列（24列目）
            
            If Left(cellValue, 1) = """" And Right(cellValue, 1) = """" Then
                cellValue = Mid(cellValue, 2, Len(cellValue) - 2)
            End If
            
            ' 登録番号をA列に記録
            If cellValue <> "" Then
                ws.Cells(rowCount, 1).Value = cellValue
                rowCount = rowCount + 1
            End If
        End If
    Loop
    
    Close #fileNum
    
    Set ReadCSVToSheet = ws
    Exit Function
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & " (エラーコード: " & Err.Number & ")", vbCritical
    If fileNum > 0 Then
        Close #fileNum
    End If
    Set ReadCSVToSheet = Nothing
End Function

' 結果ファイルの作成
' 簡易版の結果ファイル作成関数
Public Sub CreateSimpleResultFile(resultFilePath As String, xlsWs As Object, lastRow As Long, matchStatus() As Boolean, matchedRegNums() As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim resultFile As Object
    Set resultFile = fso.CreateTextFile(resultFilePath, True, False)  ' True=上書き, False=ASCII
    
    ' BOMの書き込み (UTF-8 BOMの場合)
    resultFile.Write Chr(239) & Chr(187) & Chr(191)
    
    ' ヘッダー行の書き込み
    resultFile.WriteLine "登録番号,L列データ,M列データ,マッチング状態"
    
    ' XLSX行ごとに順番に処理（オリジナルの順序を維持）
    Dim i As Long
    For i = 2 To lastRow
        ' L列とM列のデータを取得
        Dim lValue As String, mValue As String
        lValue = Trim(CStr(xlsWs.Cells(i, 12).Value))
        mValue = Trim(CStr(xlsWs.Cells(i, 13).Value))
        
        ' マッチング状態に基づいて出力
        If matchStatus(i) Then
            ' マッチした行
            resultFile.WriteLine matchedRegNums(i) & "," & lValue & "," & mValue & ",マッチング"
        Else
            ' マッチしなかった行 - 登録番号は空白
            resultFile.WriteLine "," & lValue & "," & mValue & ",未マッチング"
        End If
    Next i
    
    resultFile.Close
End Sub


'====================================================================
' 5. 処理モード別関数
'====================================================================

' 標準モードの処理関数 (最終改良版)
Public Sub ProcessFilesForStandard(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' 変数宣言
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    Dim csvDataSheet As Worksheet
    
    ' グローバル変数に値をセット
    g_csvFilePath = csvFilePath
    g_xlsxFilePath = xlsxFilePath
    g_resultFilePath = resultFilePath
    
    ' パターン値も抽出しておく
    Call ExtractPatternFromFilename(xlsxFilePath)

    ' 進捗状況を表示
    Application.StatusBar = "標準モードでファイルを処理中..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual  ' 計算を手動に設定して高速化
    
    ' 処理開始時間
    Dim startTime As Double
    startTime = Timer
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む (シートに)
    Set csvDataSheet = ReadCSVToSheet(csvFilePath)
    
    If csvDataSheet Is Nothing Then
        MsgBox "CSVファイルの読み込みに失敗しました。", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' ヘッダー追加
    csvDataSheet.Cells(1, 2).Value = "有効な登録番号"
    csvDataSheet.Cells(1, 3).Value = "aValue"
    csvDataSheet.Cells(1, 4).Value = "bValue"
    csvDataSheet.Cells(1, 5).Value = "fValue"
    csvDataSheet.Cells(1, 6).Value = "gValue"
    
    ' CSVデータを事前分析して有効な登録番号だけを抽出
    Dim j As Long, validCount As Long
    validCount = 0
    
    Dim csvLastRow As Long
    csvLastRow = csvDataSheet.Cells(csvDataSheet.Rows.Count, 1).End(xlUp).Row
    
    ' CSVの有効な登録番号を抽出するための進捗表示
    Application.StatusBar = "登録番号を抽出中... (0%)"
    
    ' 有効な登録番号をB列に抽出し、分解したデータをC～F列に格納
    For j = 2 To csvLastRow
        ' 10%ごとに進捗状況を更新
        If j Mod Int((csvLastRow - 1) / 10 + 0.5) = 0 Or j = 2 Or j = csvLastRow Then
            Application.StatusBar = "登録番号を抽出中... (" & Int((j - 2) / (csvLastRow - 2) * 100) & "%)"
            DoEvents
        End If
        
        Dim currentRegNum As String
        currentRegNum = Trim(CStr(csvDataSheet.Cells(j, 1).Value))
        
        ' 登録番号から部分を抽出
        Dim regParts As Variant
        regParts = ExtractRegistrationParts(currentRegNum)
        
        ' 有効な登録番号のみ処理
        If CBool(regParts(4)) Then
            validCount = validCount + 1
            
            ' B列に有効な登録番号をコピー
            csvDataSheet.Cells(validCount + 1, 2).Value = currentRegNum
            
            ' C～F列に分解した値を格納
            csvDataSheet.Cells(validCount + 1, 3).Value = regParts(0) ' aValue
            csvDataSheet.Cells(validCount + 1, 4).Value = regParts(1) ' bValue
            csvDataSheet.Cells(validCount + 1, 5).Value = regParts(2) ' fValue
            csvDataSheet.Cells(validCount + 1, 6).Value = regParts(3) ' gValue
        End If
    Next j
    
    ' 有効なデータがない場合は終了
    If validCount = 0 Then
        MsgBox "有効な登録番号が見つかりませんでした。", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' B列の範囲にフィルターを設定 (確認用)
    csvDataSheet.Range("B1:F" & (validCount + 1)).AutoFilter
    
    ' CSVデータを正規化（オプション - より効率的なマッチングのため）
    Application.StatusBar = "CSVデータを正規化中..."
    NormalizeCSVData csvDataSheet, 2, validCount + 1
    
    ' XLSXファイルを開く
    Application.StatusBar = "XLSXファイルを開いています..."
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    On Error Resume Next
    Set xlsWb = xlsApp.Workbooks.Open(xlsxFilePath)
    
    If Err.Number <> 0 Then
        MsgBox "XLSXファイルを開けませんでした: " & Err.Description, vbCritical
        On Error GoTo ErrorHandler
        GoTo CleanupAndExit
    End If
    
    On Error GoTo ErrorHandler
    
    Set xlsWs = xlsWb.Worksheets(1)  ' 最初のシートを使用
    
    ' データ行数を取得
    Dim lastRow As Long
    lastRow = xlsWs.Cells(xlsWs.Rows.Count, "A").End(xlUp).Row
    
    
    
   ' マッチング結果を保存する配列 - 各行の登録番号を保存
    Dim matchedRegNums() As String
    ReDim matchedRegNums(2 To lastRow)
    
    ' マッチングステータスを記録する配列
    Dim matchStatus() As Boolean
    ReDim matchStatus(2 To lastRow)
    
    ' デフォルトではすべて未マッチング
    Dim i As Long
    For i = 2 To lastRow
        matchStatus(i) = False
        matchedRegNums(i) = ""  ' 空の登録番号で初期化
    Next i
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' XLSX行ごとにCSVのデータと突き合わせ
    Application.StatusBar = "マッチング処理中... (0%)"
    
    For i = 2 To lastRow  ' ヘッダーをスキップ
        ' 進捗状況を更新（10%ごと）
        If i Mod Int((lastRow - 1) / 10 + 0.5) = 0 Or i = 2 Or i = lastRow Then
            Application.StatusBar = "マッチング処理中... (" & Int((i - 2) / (lastRow - 2) * 100) & "%)"
            DoEvents
        End If
        
        ' XLSXデータ取得...
        Dim aValue As String, bValue As String, fValue As String, gValue As String
        Dim lValue As String, mValue As String
        
        ' データ取得と変換処理...
        aValue = Trim(CStr(xlsWs.Cells(i, 1).Value))
        
        ' B列の値を変換
        Dim origBValue As String
        origBValue = Trim(CStr(xlsWs.Cells(i, 2).Value))
        bValue = GetMappedValue(origBValue, "A", "B", 2, 10)
        
        ' デフォルト処理...
        If bValue = origBValue Then
            Select Case origBValue
                Case "a"
                    bValue = "01"
                Case "b"
                    bValue = "02"
                Case "c"
                    bValue = "03"
            End Select
        End If
        
        ' 空白や特殊文字の処理
        ' aValue（4桁の数字）
        aValue = Replace(aValue, " ", "")
        If IsNumeric(aValue) Then
            ' 数値の場合、4桁になるようにゼロ埋め
            aValue = Right("0000" & aValue, 4)
        End If
        
        ' bValue（2桁の数字）
        bValue = Replace(bValue, " ", "")
        If IsNumeric(bValue) Then
            ' 数値の場合、2桁になるようにゼロ埋め
            bValue = Right("00" & bValue, 2)
        End If
        
        ' F列とG列の処理 - トリミング
        Dim origFValue As String
        origFValue = Trim(CStr(xlsWs.Cells(i, 6).Value))
        fValue = origFValue   ' マッピング関数がない場合は直接値を使用
        
        gValue = Trim(CStr(xlsWs.Cells(i, 7).Value))
        lValue = Trim(CStr(xlsWs.Cells(i, 12).Value))
        mValue = Trim(CStr(xlsWs.Cells(i, 13).Value))
        
        ' B列に抽出された有効な登録番号とマッチング
        Dim k As Long
        Dim matchFound As Boolean
        matchFound = False
        
        ' CSV行ごとにマッチングを試行
        For k = 2 To validCount + 1
            ' シートから直接値を読み取り、トリミングを追加
            Dim cellAValue As String, cellBValue As String, cellFValue As String, cellGValue As String
            
            cellAValue = Trim(CStr(csvDataSheet.Cells(k, 3).Value)) ' C列 = aValue
            cellBValue = Trim(CStr(csvDataSheet.Cells(k, 4).Value)) ' D列 = bValue 
            cellFValue = Trim(CStr(csvDataSheet.Cells(k, 5).Value)) ' E列 = fValue
            cellGValue = Trim(CStr(csvDataSheet.Cells(k, 6).Value)) ' F列 = gValue
            
            ' 空文字列の標準化
            If cellAValue = "" Then cellAValue = "0000"
            If cellBValue = "" Then cellBValue = "00"
            
            If aValue = "" Then aValue = "0000"
            If bValue = "" Then bValue = "00"
            
            ' モードによる比較ロジック
            Dim isMatch As Boolean
            isMatch = False
            
            isMatch = (cellAValue = aValue) And _
                      (cellBValue = bValue) And _
                      (cellFValue = fValue) And _
                      (cellGValue = gValue)
            
            If isMatch Then
                ' 一致する場合、その行に登録番号を記録
                Dim matchedRegNum As String
                matchedRegNum = CStr(csvDataSheet.Cells(k, 2).Value)
                
                matchedRegNums(i) = matchedRegNum  ' この行の登録番号を保存
                matchStatus(i) = True  ' マッチング状態を更新
                matchCount = matchCount + 1
                matchFound = True
                Exit For  ' 最初のマッチだけ処理
            End If
        Next k
    Next i
    
    ' 結果ファイルを作成 - シンプルなアプローチ
    Application.StatusBar = "結果ファイルを作成中..."
    CreateSimpleResultFile resultFilePath, xlsWs, lastRow, matchStatus, matchedRegNums
    ' リソース解放
    ThoroughCleanup xlsApp, xlsWb, xlsWs
    
    ' 処理結果の表示
    MsgBox "標準モードでの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "処理時間: " & Format(processingTime, "0.00") & " 秒" & vbCrLf & _
           "結果は " & resultFilePath & " に保存されました。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラーコード: " & Err.Number, vbCritical
    
CleanupAndExit:
    ' リソース解放
    ThoroughCleanup xlsApp, xlsWb, xlsWs
    
    ' 処理状態をリセット
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub

'====================================================================
' 6. メイン実行関数
'====================================================================

' ワンクリックで全処理を実行する関数
Sub OneClickProcess()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "ファイル選択中..."
    
    ' CSVファイル選択
    SelectCSVFile
    If g_csvFilePath = "" Then
        ws.Range(CELL_STATUS).Value = "キャンセルされました"
        Exit Sub
    End If
    
    ' XLSXファイル選択
    SelectXLSXFile
    If g_xlsxFilePath = "" Then
        ws.Range(CELL_STATUS).Value = "キャンセルされました"
        Exit Sub
    End If
    
    ' XLSXファイル名の条件チェック
    If Not IsValidXlsxFilename(g_xlsxFilePath) Then
        MsgBox "XLSXファイル名が条件を満たしていません。" & vbCrLf & _
               "ファイル名には以下のいずれかが含まれる必要があります：" & vbCrLf & _
               "・「データ」" & vbCrLf & _
               "・「処理」" & vbCrLf & _
               "・「4桁の数字-1桁の数字」のパターン (例: 1234-5)", vbExclamation
        ws.Range(CELL_STATUS).Value = "ファイル名エラー"
        Exit Sub
    End If
    
    ' 結果ファイルパスの自動生成（ダイアログなし）
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' XLSXファイル名からパターンを抽出
    Dim pattern As String
    pattern = ExtractPatternFromFilename(g_xlsxFilePath)
    
    ' 結果ファイル名の設定
    Dim resultFolder As String
    resultFolder = ThisWorkbook.Path
    
    ' パターンがある場合はファイル名に含める
    If pattern <> "" Then
        g_resultFilePath = fso.BuildPath(resultFolder, "結果_" & pattern & ".csv")
    Else
        g_resultFilePath = fso.BuildPath(resultFolder, "結果.csv")
    End If
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "処理中..."
    
    ' キーワードに基づいて処理を実行
    ProcessFilesForStandard g_csvFilePath, g_xlsxFilePath, g_resultFilePath
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "完了"
End Sub

' 処理終了時に徹底的にリソースをクリーンアップする関数
Public Sub ThoroughCleanup(Optional xlsApp As Object = Nothing,
 Optional xlsWb As Object = Nothing, Optional xlsWs As Object = Nothing)
    On Error Resume Next
    
    ' XLSXオブジェクトの解放
    If Not xlsWs Is Nothing Then
        Set xlsWs = Nothing
    End If
    
    If Not xlsWb Is Nothing Then
        ' ブックが開いている場合は閉じる
        On Error Resume Next
        xlsWb.Close False
        On Error GoTo 0
        Set xlsWb = Nothing
    End If
    
    If Not xlsApp Is Nothing Then
        ' アプリケーションが実行中なら終了
        On Error Resume Next
        xlsApp.Quit
        On Error GoTo 0
        Set xlsApp = Nothing
    End If
    
    ' CSVデータのクリーンアップ
    CleanupCSVData
    
    ' Excel設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' オブジェクト参照をクリア
    Set fso = Nothing  ' FileSystemObjectなど、他のオブジェクト参照があれば解放
    
    ' メモリ解放を促進
    Call CollectGarbage
    
    On Error GoTo 0
End Sub

' CSVデータを解放し、一時シートをクリーンアップする関数
Public Sub CleanupCSVData()
    On Error Resume Next
    
    ' CSVデータシートの存在確認
    Dim csvSheet As Worksheet
    Dim wsExists As Boolean
    wsExists = False
    
    For Each csvSheet In ThisWorkbook.Worksheets
        If csvSheet.Name = "CSVデータ" Then
            wsExists = True
            Exit For
        End If
    Next csvSheet
    
    ' シートが存在する場合は削除
    If wsExists Then
        Application.DisplayAlerts = False
        csvSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' マッチング分析シートがあれば削除
    wsExists = False
    
    For Each csvSheet In ThisWorkbook.Worksheets
        If csvSheet.Name = "マッチング分析" Then
            wsExists = True
            Exit For
        End If
    Next csvSheet
    
    If wsExists Then
        Application.DisplayAlerts = False
        csvSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' メモリ解放
    Set csvSheet = Nothing
    
    ' グローバル変数のクリア
    g_csvFilePath = ""
    g_xlsxFilePath = ""
    g_resultFilePath = ""
    g_fourDigits = ""
    g_oneDigit = ""
    
    ' ガベージコレクション実行（オプション）
    Call CollectGarbage
    
    On Error GoTo 0
End Sub

' ガベージコレクションを強制実行する関数 (VBA最適化用)
Private Sub CollectGarbage()
    Dim i As Long
    For i = 1 To 10
        DoEvents
        Sleep 5  ' 5ミリ秒待機
    Next i
End Sub

' Sleep関数の宣言 (Windows APIを使用)
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ProcessFilesForStandard 関数のCleanupAndExitラベル近くに追加
CleanupAndExit:
    ' CSVデータリソースの解放
    CleanupCSVData
    
    ' 画面更新と状態バーの復元
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub