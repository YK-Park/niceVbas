'환산표 250326-01
Option Explicit

' ワークシートセルの位置定義
Const CELL_PREFIX1 = "B2"        ' 接頭辞1ドロップダウン
Const CELL_PREFIX2 = "D2"        ' 接頭辞2ドロップダウン
Const CELL_CSV_PATH = "B4"       ' CSVファイルパス
Const CELL_XLSX_PATH = "B6"      ' XLSXファイルパス
Const CELL_RESULT_PATH = "B8"    ' 結果ファイルパス
Const CELL_CUSTOM_DATA1 = "B10"  ' カスタムデータ1
Const CELL_CUSTOM_DATA2 = "B12"  ' カスタムデータ2
Const CELL_STATUS = "B14"        ' 処理ステータス

' ワークシート初期化
Sub InitializeWorksheet()
    ' シートのクリア
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' シートのタイトルとラベル設定
    ws.Range("A1").Value = "CSV/XLSX データ処理ツール"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ws.Range("A2").Value = "接頭辞1:"
    ws.Range("C2").Value = "接頭辞2:"
    ws.Range("A4").Value = "CSVファイル:"
    ws.Range("A6").Value = "XLSXファイル:"
    ws.Range("A8").Value = "結果ファイル:"
    ws.Range("A10").Value = "カスタムデータ1:"
    ws.Range("A12").Value = "カスタムデータ2:"
    ws.Range("A14").Value = "処理状態:"
    
    ' デフォルト値の設定
    ws.Range(CELL_RESULT_PATH).Value = ThisWorkbook.Path & "\結果.csv"
    
    ' 注: ドロップダウンとボタンは手動で設定済み
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
            ThisWorkbook.ActiveSheet.Range(CELL_CSV_PATH).Value = .SelectedItems(1)
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
            ThisWorkbook.ActiveSheet.Range(CELL_XLSX_PATH).Value = .SelectedItems(1)
        End If
    End With
End Sub

' 結果ファイル選択
Sub SelectResultFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    
    With fd
        .Title = "結果ファイルの保存先を選択してください"
        .Filters.Clear
        .Filters.Add "CSVファイル", "*.csv"
        .InitialFileName = ThisWorkbook.Path & "\結果.csv"
        
        If .Show = -1 Then
            ThisWorkbook.ActiveSheet.Range(CELL_RESULT_PATH).Value = .SelectedItems(1)
        End If
    End With
End Sub

' 処理実行
Sub ExecuteProcess()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' 入力チェック
    If ws.Range(CELL_CSV_PATH).Value = "" Then
        MsgBox "CSVファイルを選択してください。", vbExclamation
        Exit Sub
    End If
    
    If ws.Range(CELL_XLSX_PATH).Value = "" Then
        MsgBox "XLSXファイルを選択してください。", vbExclamation
        Exit Sub
    End If
    
    If ws.Range(CELL_PREFIX1).Value = "" Then
        MsgBox "接頭辞1を選択してください。", vbExclamation
        Exit Sub
    End If
    
    If ws.Range(CELL_PREFIX2).Value = "" Then
        MsgBox "接頭辞2を選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' パスの取得
    Dim csvFilePath As String
    Dim xlsxFilePath As String
    Dim resultFilePath As String
    Dim prefixOneCode As String
    Dim prefixTwoCode As String
    Dim prefixOne As String
    Dim prefixTwo As String
    Dim customData1 As String
    Dim customData2 As String
    
    csvFilePath = ws.Range(CELL_CSV_PATH).Value
    xlsxFilePath = ws.Range(CELL_XLSX_PATH).Value
    resultFilePath = ws.Range(CELL_RESULT_PATH).Value
    
    ' 選択されたコード値とカスタムデータ
    prefixOneCode = ws.Range(CELL_PREFIX1).Value
    prefixTwoCode = ws.Range(CELL_PREFIX2).Value
    customData1 = ws.Range(CELL_CUSTOM_DATA1).Value
    customData2 = ws.Range(CELL_CUSTOM_DATA2).Value
    
    ' コードから実際の値に変換（直接マッピング）
    If IsNumeric(prefixOneCode) Then
        prefixOne = GetPrefix1Value(prefixOneCode)
    Else
        prefixOne = prefixOneCode
    End If
    
    If IsNumeric(prefixTwoCode) Then
        prefixTwo = GetPrefix2Value(prefixTwoCode)
    Else
        prefixTwo = prefixTwoCode
    End If
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "処理中..."
    
    ' 処理実行
    ProcessFiles csvFilePath, xlsxFilePath, resultFilePath, prefixOne, prefixTwo, customData1, customData2
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "完了"
End Sub

' 接頭辞1の簡単なマッピング関数
Private Function GetPrefix1Value(codeValue As String) As String
    ' コード値から接頭辞1の実際の値に変換する
    Select Case codeValue
        Case "1"
            GetPrefix1Value = "ABC"
        Case "2"
            GetPrefix1Value = "DEF"
        Case Else
            ' マッチしない場合は元の値を返す
            GetPrefix1Value = codeValue
    End Select
End Function

' 接頭辞2の簡単なマッピング関数
Private Function GetPrefix2Value(codeValue As String) As String
    ' コード値から接頭辞2の実際の値に変換する
    Select Case codeValue
        Case "1"
            GetPrefix2Value = "A"
        Case "2"
            GetPrefix2Value = "B"
        Case "3"
            GetPrefix2Value = "C"
        Case "4"
            GetPrefix2Value = "D"
        Case Else
            ' マッチしない場合は元の値を返す
            GetPrefix2Value = codeValue
    End Select
End Function

Private Sub ProcessFiles(csvFilePath As String, xlsxFilePath As String, resultFilePath As String, _
                         prefixOne As String, prefixTwo As String, customData1 As String, customData2 As String)
    ' ファイルの処理を実行
    
    ' 進捗状況を表示
    Application.StatusBar = "ファイルを処理中..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' CSVファイルを読み込む
    Dim csvData As Object
    Set csvData = CreateObject("Scripting.Dictionary")
    
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
            GoTo CleanupAndExit
        End If
    Else
        MsgBox "CSVファイルが空です。", vbExclamation
        GoTo CleanupAndExit
    End If
    
    ' CSVファイルの残りの行を読む
    Dim registrationNumbers As Object
    Set registrationNumbers = CreateObject("Scripting.Dictionary")
    
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
    
    ' XLSXファイルを開く
    Dim xlsApp As Object
    Dim xlsWb As Object
    Dim xlsWs As Object
    
    Set xlsApp = CreateObject("Excel.Application")
    xlsApp.Visible = False
    xlsApp.DisplayAlerts = False
    
    Set xlsWb = xlsApp.Workbooks.Open(xlsxFilePath)
    Set xlsWs = xlsWb.Worksheets(1)  ' 最初のシートを使用
    
    ' ヘッダー情報を取得
    Dim headerInfo As Object
    Set headerInfo = ExtractHeaderInfo(xlsWs)
    
    ' データ行数を取得
    Dim lastRow As Long
    lastRow = xlsWs.Cells(xlsWs.Rows.Count, "A").End(xlUp).Row
    
    ' 結果CSVファイルを作成
    Dim resultFile As Object
    Set resultFile = fso.CreateTextFile(resultFilePath, True, True)  ' True=上書き, True=Unicode
    
    ' BOMの書き込み (UTF-8 BOMの場合)
    resultFile.Write Chr(239) & Chr(187) & Chr(191)
    
    ' ヘッダー行の書き込み
    resultFile.WriteLine "登録番号,L列データ,M列データ"
    
    ' マッチングカウンター
    Dim matchCount As Long
    matchCount = 0
    
    ' XLSX行ごとにCSVのデータと突き合わせ
    For i = 2 To lastRow  ' ヘッダーをスキップ
        Dim aValue As String
        Dim bValue As String
        Dim fValue As String
        Dim gValue As String
        Dim lValue As String
        Dim mValue As String
        
        ' XLSXからデータを取得
        aValue = CStr(xlsWs.Cells(i, 1).Value)  ' A列の値
        
        ' B列の変換（a,b,c → 01,02,03）
        Select Case LCase(xlsWs.Cells(i, 2).Value)
            Case "a"
                bValue = "01"
            Case "b"
                bValue = "02"
            Case "c"
                bValue = "03"
            Case Else
                bValue = "00"  ' デフォルト値
        End Select
        
        ' F列の変換（1,2,3 → mmddWWW形式）
        Dim fKey As String
        fKey = CStr(xlsWs.Cells(i, 6).Value)  ' F列の値
        
        If headerInfo.Exists(fKey) Then
            fValue = headerInfo(fKey)
        Else
            fValue = "0000XXX"  ' デフォルト値
        End If
        
        ' G列の値
        gValue = CStr(xlsWs.Cells(i, 7).Value)  ' G列の値
        
        ' L, M列の値を取得
        lValue = CStr(xlsWs.Cells(i, 12).Value)  ' L列
        mValue = CStr(xlsWs.Cells(i, 13).Value)  ' M列
        
        ' CSVの登録番号と照合
        Dim matchFound As Boolean
        matchFound = False
        
        Dim csvRegKey As Variant
        For Each csvRegKey In registrationNumbers.Keys
            Dim currentRegNum As String
            currentRegNum = CStr(csvRegKey)
            
            ' 接頭辞チェック
            If Mid(currentRegNum, 1, Len(prefixOne)) = prefixOne And _
               Mid(currentRegNum, Len(prefixOne) + 2, 1) = prefixTwo Then
                
                ' 各部分を抽出して比較
                Dim csvAValue As String
                Dim csvBValue As String
                Dim csvFValue As String
                Dim csvGValue As String
                
                ' 登録番号から各部分を抽出（例：ABC-D01232010101FRIA）
                ' 接頭辞部分をスキップして、それぞれの部分を取得
                ' 位置はフォーマットによって調整が必要
                Dim startPos As Long
                startPos = Len(prefixOne) + 3  ' "ABC-D" の後
                
                ' それぞれの部分を抽出（位置は実際のフォーマットに合わせて調整）
                csvAValue = Mid(currentRegNum, startPos, 4)               ' 例: 0123
                csvBValue = Mid(currentRegNum, startPos + 4, 2)           ' 例: 01
                csvFValue = Mid(currentRegNum, startPos + 6, 7)           ' 例: 0101FRI
                csvGValue = Mid(currentRegNum, startPos + 13, 1)          ' 例: A
                
                ' カスタムデータも考慮（指定がある場合）
                If (aValue = csvAValue Or customData1 = csvAValue) And _
                   (bValue = csvBValue) And _
                   (fValue = csvFValue) And _
                   (gValue = csvGValue Or customData2 = csvGValue) Then
                    
                    ' マッチしたデータを結果ファイルに書き込み
                    resultFile.WriteLine currentRegNum & "," & lValue & "," & mValue
                    matchCount = matchCount + 1
                    matchFound = True
                    Exit For
                End If
            End If
        Next csvRegKey
    Next i
    
    ' ファイルを閉じる
    resultFile.Close
    xlsWb.Close False
    xlsApp.Quit
    
    Set xlsWs = Nothing
    Set xlsWb = Nothing
    Set xlsApp = Nothing
    
    ' 処理結果の表示
    MsgBox "処理が完了しました。" & vbCrLf & _
           "合計 " & (lastRow - 1) & " 件のXLSXデータから " & matchCount & " 件のマッチングを見つけました。" & vbCrLf & _
           "結果は " & resultFilePath & " に保存されました。", vbInformation
    
CleanupAndExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanupAndExit
End Sub

Private Function ExtractHeaderInfo(ws As Object) As Object
    ' F列のヘッダー情報を解析
    Dim headerDict As Object
    Set headerDict = CreateObject("Scripting.Dictionary")
    
    Dim headerText As String
    headerText = ws.Cells(1, 6).Value
    
    ' マルチラインセルの場合
    If InStr(headerText, vbLf) > 0 Then
        Dim lines() As String
        lines = Split(headerText, vbLf)
        
        Dim line As Variant
        For Each line In lines
            ParseHeaderLine CStr(line), headerDict
        Next
    Else
        ' 単一行の場合
        ParseHeaderLine headerText, headerDict
    End If
    
    Set ExtractHeaderInfo = headerDict
End Function

Private Sub ParseHeaderLine(lineText As String, dict As Object)
    ' 「1: 金曜日(1/1)」形式の行を解析
    
    Dim pos As Long
    pos = InStr(lineText, ":")
    
    If pos > 0 Then
        Dim key As String
        Dim value As String
        Dim dateInfo As String
        
        key = Trim(Left(lineText, pos - 1))
        value = Trim(Mid(lineText, pos + 1))
        
        ' 日付情報を抽出（例: 金曜日(1/1) から 0101FRI を生成）
        If InStr(value, "(") > 0 And InStr(value, ")") > 0 Then
            Dim startPos As Long
            Dim endPos As Long
            startPos = InStr(value, "(") + 1
            endPos = InStr(value, ")") - 1
            
            If startPos <= endPos Then
                dateInfo = Mid(value, startPos, endPos - startPos + 1)
                
                ' 日付の解析（1/1 形式を想定）
                If InStr(dateInfo, "/") > 0 Then
                    Dim dateParts() As String
                    dateParts = Split(dateInfo, "/")
                    
                    If UBound(dateParts) >= 1 Then
                        Dim month As String
                        Dim day As String
                        month = Format(dateParts(0), "00")
                        day = Format(dateParts(1), "00")
                        
                        ' 曜日の特定
                        Dim dayOfWeek As String
                        If InStr(value, "月曜") > 0 Then
                            dayOfWeek = "MON"
                        ElseIf InStr(value, "火曜") > 0 Then
                            dayOfWeek = "TUE"
                        ElseIf InStr(value, "水曜") > 0 Then
                            dayOfWeek = "WED"
                        ElseIf InStr(value, "木曜") > 0 Then
                            dayOfWeek = "THU"
                        ElseIf InStr(value, "金曜") > 0 Then
                            dayOfWeek = "FRI"
                        ElseIf InStr(value, "土曜") > 0 Then
                            dayOfWeek = "SAT"
                        ElseIf InStr(value, "日曜") > 0 Then
                            dayOfWeek = "SUN"
                        Else
                            dayOfWeek = "XXX"
                        End If
                        
                        ' mmddWWW形式の作成
                        dict(key) = month & day & dayOfWeek
                    End If
                End If
            End If
        End If
    End If
End Sub

' ワークブック起動時に初期化を実行
Sub Auto_Open()
    InitializeWorksheet
End Sub