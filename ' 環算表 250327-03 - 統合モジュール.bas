' 環算表 250327-03 - 統合モジュール
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

' ワークブック起動時に初期化を実行
Sub Auto_Open()
    InitializeWorksheet
End Sub

' ワークシート初期化
Sub InitializeWorksheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' シートのタイトルとラベル設定
    ws.Range("A1").Value = "CSV/XLSX データ処理ツール"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ws.Range("A10").Value = "処理状態:"
    
    ' 処理ステータスをクリア
    ws.Range(CELL_STATUS).Value = ""
End Sub

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

' XLSX ファイル名からキーワードを検出して処理タイプを決定する関数
Private Function DetectProcessingType(filename As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String
    baseName = fso.GetFileName(filename)
    
    If InStr(1, baseName, "集計") > 0 Then
        DetectProcessingType = "集計"
    ElseIf InStr(1, baseName, "分析") > 0 Then
        DetectProcessingType = "分析"
    ElseIf InStr(1, baseName, "処理") > 0 Then
        DetectProcessingType = "処理"
    Else
        DetectProcessingType = "標準"
    End If
End Function

' キーワードに基づいて適切な処理関数を呼び出す
Private Sub ProcessFilesBasedOnKeyword(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    Dim processingType As String
    processingType = DetectProcessingType(xlsxFilePath)
    
    ' ステータスの更新
    ThisWorkbook.ActiveSheet.Range(CELL_STATUS).Value = processingType & "モードで処理中..."
    
    Select Case processingType
        Case "集計"
            ProcessFilesForSyukei csvFilePath, xlsxFilePath, resultFilePath
        Case "分析"
            ProcessFilesForBunseki csvFilePath, xlsxFilePath, resultFilePath
        Case "処理"
            ProcessFilesForSyori csvFilePath, xlsxFilePath, resultFilePath
        Case Else
            ProcessFilesForStandard csvFilePath, xlsxFilePath, resultFilePath
    End Select
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

' xValueに基づいて値を変換する関数
Public Function ConvertValueIfMatch(xValue As String, comparisonValue As String, startRow As Integer, endRow As Integer) As String
    ' xValueが参照シートの1列目の値と一致する場合、2列目の値を3列目の値に変換する
    
    Dim mappingSheet As Worksheet
    Set mappingSheet = ThisWorkbook.Worksheets("参照")
    
    Dim i As Integer
    For i = startRow To endRow
        ' A列の値とxValueを比較
        If CStr(mappingSheet.Range("A" & i).Value) = CStr(xValue) Then
            ' B列の値とcomparisonValueが一致する場合、C列の値を返す
            If CStr(mappingSheet.Range("B" & i).Value) = CStr(comparisonValue) Then
                ConvertValueIfMatch = CStr(mappingSheet.Range("C" & i).Value)
                Exit Function
            End If
        End If
    Next i
    
    ' マッチしない場合は元の値を返す
    ConvertValueIfMatch = comparisonValue
End Function

' F列の値をマッピングする共通関数
Public Function GetFValueMapping(fCode As String) As String
    ' 参照シートから取得
    Dim mappedValue As String
    mappedValue = GetMappedValue(fCode, "C", "D", 2, 20)
    
    ' マッピングが見つからない場合はデフォルト値を使用
    If mappedValue = fCode Then
        mappedValue = "0000XXX"
    End If
    
    GetFValueMapping = mappedValue
End Function

' 登録番号から各部分を抽出する関数
Public Function ExtractRegistrationParts(regNum As String) As Object
    Dim parts As Object
    Set parts = CreateObject("Scripting.Dictionary")
    
    ' 固定位置から各部分を抽出
    ' 例: regNum = "ABC-D01232010101FRIA" の場合
    parts("aValue") = Mid(regNum, 6, 4)    ' "0123"
    parts("bValue") = Mid(regNum, 10, 2)   ' "01"
    parts("fValue") = Mid(regNum, 12, 7)   ' "0101FRI"
    parts("gValue") = Mid(regNum, 19, 1)   ' "A"
    parts("isValid") = True
    
    Set ExtractRegistrationParts = parts
End Function

' 登録番号の部分とXLSXデータを比較する関数
Public Function IsDataMatching(regParts As Object, aValue As String, bValue As String, _
                              fValue As String, gValue As String) As Boolean
    ' 部分が有効でない場合は一致しない
    If Not regParts("isValid") Then
        IsDataMatching = False
        Exit Function
    End If
    
    ' 各部分を比較
    IsDataMatching = (regParts("aValue") = aValue) And _
                    (regParts("bValue") = bValue) And _
                    (regParts("fValue") = fValue) And _
                    (regParts("gValue") = gValue)
End Function

'====================================================================
' 4. ファイル処理関数
'====================================================================

' CSVファイルから登録番号データを読み込む
Public Function ReadCSVFile(csvFilePath As String) As Object
    Dim registrationNumbers As Object
    Set registrationNumbers = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim csvFile As Object
    Set csvFile = fso.OpenTextFile(csvFilePath, 1, False, -1)  ' 1=ForReading
    
    Dim line As String
    Dim csvValues() As String
    Dim xIndex As Long
    
    ' CSVファイルの1行目を読んで列のインデックスを特定
    If Not csvFile.AtEndOfStream Then
        line = csvFile.ReadLine
        csvValues = Split(line, ",")
        
        ' X列（登録番号列）のインデックスを取得
        xIndex = -1
        For i = 0 To UBound(csvValues)
            If i = 23 Then  ' X列は24番目（0から始まる場合は23）
                xIndex = i
                Exit For
            End If
        Next i
        
        If xIndex = -1 Then
            MsgBox "CSVファイルにX列（登録番号列）が見つかりません。", vbExclamation
            Set ReadCSVFile = Nothing
            Exit Function
        End If
    Else
        MsgBox "CSVファイルが空です。", vbExclamation
        Set ReadCSVFile = Nothing
        Exit Function
    End If
    
    ' CSVファイルの残りの行を読む
    Do Until csvFile.AtEndOfStream
        line = csvFile.ReadLine
        csvValues = Split(line, ",")
        
        If UBound(csvValues) >= xIndex Then
            Dim regNum As String
            regNum = Trim(csvValues(xIndex))
            
            If regNum <> "" Then
                ' 同じ登録番号が複数ある場合は最後のものを使用
                registrationNumbers(regNum) = line
            End If
        End If
    Loop
    
    csvFile.Close
    
    Set ReadCSVFile = registrationNumbers
End Function

' 結果ファイルの作成
Public Sub CreateResultFile(resultFilePath As String, matchedData As Collection)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim resultFile As Object
    Set resultFile = fso.CreateTextFile(resultFilePath, True, True)  ' True=上書き, True=Unicode
    
    ' BOMの書き込み (UTF-8 BOMの場合)
    resultFile.Write Chr(239) & Chr(187) & Chr(191)
    
    ' ヘッダー行の書き込み
    resultFile.WriteLine "登録番号,L列データ,M列データ"
    
    ' マッチングデータの書き込み
    Dim item As Variant
    For Each item In matchedData
        resultFile.WriteLine item
    Next item
    
    resultFile.Close
End Sub

' エラー処理を行う
Public Sub HandleError(errorMsg As String)
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & errorMsg, vbCritical
End Sub

'====================================================================
' 5. 処理モード別関数
'====================================================================

' 標準モードの処理関数
Public Sub ProcessFilesForStandard(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' 進捗状況を表示
    Application.StatusBar = "標準モードでファイルを処理中..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む
    Dim registrationNumbers As Object
    Set registrationNumbers = ReadCSVFile(csvFilePath)
    
    If registrationNumbers Is Nothing Then
        GoTo CleanupAndExit
    End If
    
    ' XLSXファイルを開く
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    Set xlsWb = xlsApp.Workbooks.Open(xlsxFilePath)
    Set xlsWs = xlsWb.Worksheets(1)  ' 最初のシートを使用
    
    ' データ行数を取得
    Dim lastRow As Long
    lastRow = xlsWs.Cells(xlsWs.Rows.Count, "A").End(xlUp).Row
    
    ' マッチング結果を保存するコレクション
    Dim matchedResults As Collection
    Set matchedResults = New Collection
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' XLSX行ごとにCSVのデータと突き合わせ
    Dim i As Long
    For i = 2 To lastRow  ' ヘッダーをスキップ
        Dim aValue As String
        Dim bValue As String
        Dim fValue As String
        Dim gValue As String
        Dim lValue As String
        Dim mValue As String
        
        ' 元のデータを取得
        aValue = CStr(xlsWs.Cells(i, 1).Value)  ' A列の値
        
        ' B列の値を変換 - GetMappedValueを使用
        Dim origBValue As String
        origBValue = CStr(xlsWs.Cells(i, 2).Value)
        bValue = GetMappedValue(origBValue, "A", "B", 2, 10)
        
        ' マッピングが見つからない場合のデフォルト処理
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
        
        ' F列の値を取得して変換 - GetMappedValueを使用
        Dim origFValue As String
        origFValue = CStr(xlsWs.Cells(i, 6).Value)  ' F列の値
        fValue = GetFValueMapping(origFValue)
        
        ' G列の値を取得
        gValue = CStr(xlsWs.Cells(i, 7).Value)  ' G列の値
        
        ' L, M列の値を取得
        lValue = CStr(xlsWs.Cells(i, 12).Value)  ' L列
        mValue = CStr(xlsWs.Cells(i, 13).Value)  ' M列
        
        ' CSVの登録番号を検索してマッチング
        Dim csvRegKey As Variant
        Dim matchFound As Boolean
        matchFound = False
        
        For Each csvRegKey In registrationNumbers.Keys
            Dim currentRegNum As String
            currentRegNum = CStr(csvRegKey)
            
            ' 登録番号から部分を抽出
            Dim regParts As Object
            Set regParts = ExtractRegistrationParts(currentRegNum)
            
            ' 抽出された部分とXLSXデータを比較
            If IsDataMatching(regParts, aValue, bValue, fValue, gValue) Then
                ' 一致する場合、結果に追加
                matchedResults.Add currentRegNum & "," & lValue & "," & mValue
                matchCount = matchCount + 1
                matchFound = True
                Exit For
            End If
        Next csvRegKey
    Next i
    
    ' 結果ファイルを作成
    CreateResultFile resultFilePath, matchedResults
    
    ' ファイルを閉じる
    xlsWb.Close False
    xlsApp.Quit
    
    Set xlsWs = Nothing
    Set xlsWb = Nothing
    Set xlsApp = Nothing
    
    ' 処理結果の表示
    MsgBox "標準モードでの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "結果は " & resultFilePath & " に保存されました。", vbInformation
    
CleanupAndExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    HandleError Err.Description
    Resume CleanupAndExit
End Sub

' 集計モード用の処理関数
Public Sub ProcessFilesForSyukei(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' 進捗状況を表示
    Application.StatusBar = "集計モードでファイルを処理中..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む
    Dim registrationNumbers As Object
    Set registrationNumbers = ReadCSVFile(csvFilePath)
    
    If registrationNumbers Is Nothing Then
        GoTo CleanupAndExit
    End If
    
    ' XLSXファイルを開く
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    Set xlsWb = xlsApp.Workbooks.Open(xlsxFilePath)
    Set xlsWs = xlsWb.Worksheets(1)  ' 最初のシートを使用
    
    ' データ行数を取得
    Dim lastRow As Long
    lastRow = xlsWs.Cells(xlsWs.Rows.Count, "A").End(xlUp).Row
    
    ' マッチング結果を保存するコレクション
    Dim matchedResults As Collection
    Set matchedResults = New Collection
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' XLSX行ごとにCSVのデータと突き合わせ
    Dim i As Long
    For i = 2 To lastRow  ' ヘッダーをスキップ
        Dim aValue As String
        Dim bValue As String
        Dim fValue As String
        Dim gValue As String
        Dim lValue As String
        Dim mValue As String
        
        ' 集計モード特有の処理: A列の値を常に4桁に整形
        Dim origAValue As String
        origAValue = CStr(xlsWs.Cells(i, 1).Value)
        aValue = Format(Val(origAValue), "0000")
        
        ' B列は常に "01" に設定（集計モード特有の処理）
        bValue = "01"
        
        ' F列の値を取得して変換 - GetMappedValueを使用
        Dim origFValue As String
        origFValue = CStr(xlsWs.Cells(i, 6).Value)  ' F列の値
        fValue = GetFValueMapping(origFValue)
        
        ' G列の値を取得
        gValue = CStr(xlsWs.Cells(i, 7).Value)  ' G列の値
        
        ' L, M列の値を取得
        lValue = CStr(xlsWs.Cells(i, 12).Value)  ' L列
        mValue = CStr(xlsWs.Cells(i, 13).Value)  ' M列
        
        ' CSVの登録番号を検索してマッチング
        Dim csvRegKey As Variant
        Dim matchFound As Boolean
        matchFound = False
        
        For Each csvRegKey In registrationNumbers.Keys
            Dim currentRegNum As String
            currentRegNum = CStr(csvRegKey)
            
            ' 登録番号から部分を抽出
            Dim regParts As Object
            Set regParts = ExtractRegistrationParts(currentRegNum)
            
            ' 抽出された部分とXLSXデータを比較
            If IsDataMatching(regParts, aValue, bValue, fValue, gValue) Then
                ' 一致する場合、結果に追加
                matchedResults.Add currentRegNum & "," & lValue & "," & mValue
                matchCount = matchCount + 1
                matchFound = True
                Exit For
            End If
        Next csvRegKey
    Next i
    
    ' 結果ファイルを作成
    CreateResultFile resultFilePath, matchedResults
    
    ' ファイルを閉じる
    xlsWb.Close False
    xlsApp.Quit
    
    Set xlsWs = Nothing
    Set xlsWb = Nothing
    Set xlsApp = Nothing
    
    ' 処理結果の表示
    MsgBox "集計モードでの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "結果は " & resultFilePath & " に保存されました。", vbInformation
    
CleanupAndExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    HandleError Err.Description
    Resume CleanupAndExit
End Sub

' 分析モード用の処理関数 - 二つのファイルを処理してfValue基準で並べ替え
Public Sub ProcessFilesForBunseki(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' 進捗状況を表示
    Application.StatusBar = "分析モードでファイルを処理中..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む
    Dim registrationNumbers As Object
    Set registrationNumbers = ReadCSVFile(csvFilePath)
    
    If registrationNumbers Is Nothing Then
        GoTo CleanupAndExit
    End If
    
    ' 2つのXLSXファイル名を取得
    Dim xlsxPath1 As String
    Dim xlsxPath2 As String
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "2つ目のXLSXファイルを選択してください"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            xlsxPath1 = xlsxFilePath    ' 1つ目のファイル
            xlsxPath2 = .SelectedItems(1) ' 2つ目のファイル
        Else
            MsgBox "2つ目のファイルが選択されていません。処理を中止します。", vbExclamation
            GoTo CleanupAndExit
        End If
    End With
    
    ' XLSXファイルを開く
    Dim xlsApp As Object
    Dim xlsWb1 As Object, xlsWs1 As Object
    Dim xlsWb2 As Object, xlsWs2 As Object
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    Set xlsWb1 = xlsApp.Workbooks.Open(xlsxPath1)
    Set xlsWs1 = xlsWb1.Worksheets(1)
    
    Set xlsWb2 = xlsApp.Workbooks.Open(xlsxPath2)
    Set xlsWs2 = xlsWb2.Worksheets(1)
    
    ' データ行数を取得
    Dim lastRow1 As Long, lastRow2 As Long
    lastRow1 = xlsWs1.Cells(xlsWs1.Rows.Count, "A").End(xlUp).Row
    lastRow2 = xlsWs2.Cells(xlsWs2.Rows.Count, "A").End(xlUp).Row
    
    ' マッチング結果を保存する配列
    Dim matchResults() As String
    ReDim matchResults(1 To 10000, 1 To 2)  ' 1列目: fValue, 2列目: 結果文字列
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' 両方のファイルを処理
    ProcessXlsxFile xlsWs1, lastRow1, xlsWs2, lastRow2, registrationNumbers, matchResults, matchCount
    ProcessXlsxFile xlsWs2, lastRow2, xlsWs1, lastRow1, registrationNumbers, matchResults, matchCount
    
    ' 結果配列のサイズを調整
    ReDim Preserve matchResults(1 To matchCount, 1 To 2)
    
    ' fValueで並べ替え
    SortMatchResults matchResults, matchCount
    
    ' ソートされた結果をCollectionに変換
    Dim sortedResults As Collection
    Set sortedResults = New Collection
    
    Dim k As Long
    For k = 1 To matchCount
        sortedResults.Add matchResults(k, 2)
    Next k
    
    ' 結果ファイルを作成
    CreateResultFile resultFilePath, sortedResults
    
    ' ファイルを閉じる
    xlsWb1.Close False
    xlsWb2.Close False
    xlsApp.Quit
    
    ' メモリ解放
    Set xlsWs1 = Nothing
    Set xlsWs2 = Nothing
    Set xlsWb1 = Nothing
    Set xlsWb2 = Nothing
    Set xlsApp = Nothing
    
    ' 処理結果の表示
    MsgBox "分析モードでの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow1 - 1 + lastRow2 - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "結果は " & resultFilePath & " に保存され、fValue順にソートされています。", vbInformation
    
CleanupAndExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    HandleError Err.Description
    Resume CleanupAndExit
End Sub

' XLSXファイル処理関数（再利用可能）
Private Sub ProcessXlsxFile(ws As Object, lastRow As Long, otherWs As Object, otherLastRow As Long, _
                           registrationNumbers As Object, ByRef matchResults() As String, ByRef matchCount As Long)
    Dim i As Long, j As Long
    
    For i = 2 To lastRow  ' ヘッダーをスキップ
        ' 基本データの取得
        Dim aValue As String, bValue As String, fValue As String, gValue As String
        Dim lValue As String, mValue As String
        Dim origBValue As String, origFValue As String
        
        aValue = CStr(ws.Cells(i, 1).Value)
        
        ' B列の処理
        origBValue = CStr(ws.Cells(i, 2).Value)
        bValue = GetMappedValue(origBValue, "A", "B", 2, 10)
        
        ' F列の処理
        origFValue = CStr(ws.Cells(i, 6).Value)
        fValue = GetMappedValue(origFValue, "C", "D", 2, 20)
        
        ' F列のデフォルト値が必要な場合
        If fValue = origFValue Then
            fValue = "0101XXX"  ' 分析モード特有のデフォルト値
        End If
        
        ' G列の処理 - 大文字に統一
        gValue = UCase(CStr(ws.Cells(i, 7).Value))
        
        ' 他のファイルから対応データを検索
        Dim found As Boolean
        found = False
        
        For j = 2 To otherLastRow
            If CStr(otherWs.Cells(j, 1).Value) = aValue Then
                lValue = CStr(otherWs.Cells(j, 12).Value)
                mValue = CStr(otherWs.Cells(j, 13).Value)
                found = True
                Exit For
            End If
        Next j
        
        ' 他のファイルに対応するデータがない場合は現在のファイルから取得
        If Not found Then
            lValue = CStr(ws.Cells(i, 12).Value)
            mValue = CStr(ws.Cells(i, 13).Value)
        End If
        
        ' CSV登録番号の検索とマッチング
        Dim csvRegKey As Variant
        Dim matchFound As Boolean
        matchFound = False
        
        For Each csvRegKey In registrationNumbers.Keys
            Dim currentRegNum As String
            currentRegNum = CStr(csvRegKey)
            
            ' 登録番号から部分を抽出
            Dim regParts As Object
            Set regParts = ExtractRegistrationParts(currentRegNum)
            
            ' 抽出された部分とXLSXデータを比較
            If IsDataMatching(regParts, aValue, bValue, fValue, gValue) Then
                ' 一致する場合、結果に追加
                matchCount = matchCount + 1
                matchResults(matchCount, 1) = fValue
                matchResults(matchCount, 2) = currentRegNum & "," & lValue & "," & mValue
                matchFound = True
                Exit For
            End If
        Next csvRegKey
    Next i
End Sub

' 結果をfValueでソートする関数
Private Sub SortMatchResults(ByRef matchResults() As String, ByVal matchCount As Long)
    Dim temp1 As String, temp2 As String
    Dim k As Long, l As Long
    
    For k = 1 To matchCount - 1
        For l = k + 1 To matchCount
            If matchResults(k, 1) > matchResults(l, 1) Then
                ' fValue部分を交換
                temp1 = matchResults(k, 1)
                matchResults(k, 1) = matchResults(l, 1)
                matchResults(l, 1) = temp1
                
                ' 結果部分を交換
                temp2 = matchResults(k, 2)
                matchResults(k, 2) = matchResults(l, 2)
                matchResults(l, 2) = temp2
            End If
        Next l
    Next k
End Sub

' 処理モード用の処理関数
Public Sub ProcessFilesForSyori(csvFilePath As String, xlsxFilePath As String, resultFilePath As String)
    ' 進捗状況を表示
    Application.StatusBar = "処理モードでファイルを処理中..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む
    Dim registrationNumbers As Object
    Set registrationNumbers = ReadCSVFile(csvFilePath)
    
    If registrationNumbers Is Nothing Then
        GoTo CleanupAndExit
    End If
    
    ' XLSXファイルを開く
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    Set xlsWb = xlsApp.Workbooks.Open(xlsxFilePath)
    Set xlsWs = xlsWb.Worksheets(1)  ' 最初のシートを使用
    
    ' データ行数を取得
    Dim lastRow As Long
    lastRow = xlsWs.Cells(xlsWs.Rows.Count, "A").End(xlUp).Row
    
    ' マッチング結果を保存するコレクション
    Dim matchedResults As Collection
    Set matchedResults = New Collection
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' XLSX行ごとにCSVのデータと突き合わせ
    Dim i As Long
    For i = 2 To lastRow  ' ヘッダーをスキップ
        Dim aValue As String
        Dim bValue As String
        Dim fValue As String
        Dim gValue As String
        Dim lValue As String
        Dim mValue As String
        
        ' 処理モード特有の処理
        ' ファイル名からパターンを取得して適用
        If g_fourDigits <> "" Then
            aValue = g_fourDigits
        Else
            aValue = CStr(xlsWs.Cells(i, 1).Value)
        End If
        
        ' B列の処理
        Dim origBValue As String
        origBValue = CStr(xlsWs.Cells(i, 2).Value)
        
        ' 特定のパターンがあれば適用
        If g_oneDigit <> "" Then
            Select Case g_oneDigit
                Case "1"
                    bValue = "01"
                Case "2"
                    bValue = "02"
                Case "3"
                    bValue = "03"
                Case Else
                    bValue = "0" & g_oneDigit
            End Select
        Else
            ' GetMappedValueを使用
            bValue = GetMappedValue(origBValue, "A", "B", 2, 10)
            
            ' マッピングが見つからない場合のデフォルト処理
            If bValue = origBValue Then
                Select Case origBValue
                    Case "a"
                        bValue = "01"
                    Case "b"
                        bValue = "02"
                    Case "c"
                        bValue = "03"
                    Case Else
                        bValue = "00"
                End Select
            End If
        End If
        
        ' F列の処理
        Dim origFValue As String
        origFValue = CStr(xlsWs.Cells(i, 6).Value)
        fValue = GetFValueMapping(origFValue)
        
        ' G列の処理
        gValue = CStr(xlsWs.Cells(i, 7).Value)
        
        ' L, M列のデータ取得
        lValue = CStr(xlsWs.Cells(i, 12).Value)
        mValue = CStr(xlsWs.Cells(i, 13).Value)
        
        ' CSVの登録番号を検索してマッチング
        Dim csvRegKey As Variant
        Dim matchFound As Boolean
        matchFound = False
        
        For Each csvRegKey In registrationNumbers.Keys
            Dim currentRegNum As String
            currentRegNum = CStr(csvRegKey)
            
            ' 登録番号から部分を抽出
            Dim regParts As Object
            Set regParts = ExtractRegistrationParts(currentRegNum)
            
            ' 抽出された部分とXLSXデータを比較
            If IsDataMatching(regParts, aValue, bValue, fValue, gValue) Then
                ' 一致する場合、結果に追加
                matchedResults.Add currentRegNum & "," & lValue & "," & mValue
                matchCount = matchCount + 1
                matchFound = True
                Exit For
            End If
        Next csvRegKey
    Next i
    
    ' 結果ファイルを作成
    CreateResultFile resultFilePath, matchedResults
    
    ' ファイルを閉じる
    xlsWb.Close False
    xlsApp.Quit
    
    Set xlsWs = Nothing
    Set xlsWb = Nothing
    Set xlsApp = Nothing
    
    ' 処理結果の表示
    MsgBox "処理モードでの処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "結果は " & resultFilePath & " に保存されました。", vbInformation
    
CleanupAndExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    HandleError Err.Description
    Resume CleanupAndExit
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
    ProcessFilesBasedOnKeyword g_csvFilePath, g_xlsxFilePath, g_resultFilePath
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "完了"
End Sub