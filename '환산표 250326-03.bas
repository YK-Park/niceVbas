'환산표 250326-03
' ワークブック起動時に初期化を実行
Sub Auto_Open()
    InitializeWorksheet
End SubOption Explicit

' グローバル変数 - 明示的にPrivateとして宣言
Private g_csvFilePath As String
Private g_xlsxFilePath As String
Private g_resultFilePath As String

' ワークシートセルの位置定義
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
    
    ws.Range("A10").Value = "カスタムデータ1:"
    ws.Range("A12").Value = "カスタムデータ2:"
    ws.Range("A14").Value = "処理状態:"
    
    ' デフォルト値の設定
    g_resultFilePath = ThisWorkbook.Path & "\結果.csv"
    
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

' ファイル名から4桁の数字-1桁の数字のパターンを抽出する関数
Private Function ExtractPatternFromFilename(filename As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String
    baseName = fso.GetFileName(filename)
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = "\d{4}-\d"
        .Global = True
        .IgnoreCase = True
        
        Dim matches As Object
        Set matches = .Execute(baseName)
        
        If matches.Count > 0 Then
            ExtractPatternFromFilename = matches(0)
        Else
            ExtractPatternFromFilename = ""
        End If
    End With
End Function

' ファイル名から条件値を抽出する関数
Private Function ExtractConditionsFromFilename(filename As String, ByRef condition1Value As String, ByRef condition2Value As String, _
                                               ByRef fourDigits As String, ByRef oneDigit As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String
    baseName = fso.GetFileName(filename)
    
    ' 条件1: "データ"が含まれているか
    If InStr(1, baseName, "データ") > 0 Then
        condition1Value = "データ"
    Else
        condition1Value = ""
    End If
    
    ' 条件2: "処理"が含まれているか
    If InStr(1, baseName, "処理") > 0 Then
        condition2Value = "処理"
    Else
        condition2Value = ""
    End If
    
    ' 条件3: 4桁の数字-1桁の数字のパターン
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
            fourDigits = matches(0).SubMatches(0)
            oneDigit = matches(0).SubMatches(1)
        Else
            fourDigits = ""
            oneDigit = ""
        End If
    End With
    
    ' いずれかの条件が満たされている場合はTrueを返す
    ExtractConditionsFromFilename = (condition1Value <> "" Or condition2Value <> "" Or (fourDigits <> "" And oneDigit <> ""))
End Function

' XLSXファイル名が条件を満たしているか確認する関数
Private Function IsValidXlsxFilename(filename As String) As Boolean
    Dim condition1Value As String
    Dim condition2Value As String
    Dim fourDigits As String
    Dim oneDigit As String
    
    ' 抽出関数を使用して条件を確認
    IsValidXlsxFilename = ExtractConditionsFromFilename(filename, condition1Value, condition2Value, fourDigits, oneDigit)
End Function
    
    ' カスタムデータを取得
    Dim customData1 As String
    Dim customData2 As String
    
    customData1 = ws.Range(CELL_CUSTOM_DATA1).Value
    customData2 = ws.Range(CELL_CUSTOM_DATA2).Value
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "処理中..."
    
    ' 処理実行
    ProcessFiles g_csvFilePath, g_xlsxFilePath, g_resultFilePath, customData1, customData2
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "完了"
End Sub

' マッピングテーブルから実際の値を取得する関数
Private Function GetMappedValue(codeValue As String, codeCol As String, valueCol As String, startRow As Integer, endRow As Integer) As String
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
Private Function ConvertValueIfMatch(xValue As String, comparisonValue As String, startRow As Integer, endRow As Integer) As String
    ' xValueが参照シートの1列目の値と一致する場合、2列目の値を3列目の値に変換する
    ' xValue: 比較する値（A列の値）
    ' comparisonValue: 変換対象の値（B列の値）
    ' startRow/endRow: 検索行範囲
    
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

Private Sub ProcessFiles(csvFilePath As String, xlsxFilePath As String, resultFilePath As String, _
                         customData1 As String, customData2 As String)
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
    
    ' ファイル名から条件値を抽出
    Dim condition1Value As String
    Dim condition2Value As String
    Dim fourDigits As String
    Dim oneDigit As String
    
    Dim hasConditions As Boolean
    hasConditions = ExtractConditionsFromFilename(xlsxFilePath, condition1Value, condition2Value, fourDigits, oneDigit)
    
    ' XLSXファイルのヘッダー情報を取得
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
        Dim prefixOne As String
        Dim prefixTwo As String
        
        ' ファイル名から抽出した条件値を各列に適用
        ' 元のデータを取得
        Dim origAValue As String
        Dim origBValue As String
        Dim origFValue As String
        Dim origGValue As String
        
        origAValue = CStr(xlsWs.Cells(i, 1).Value)  ' A列の元の値
        origBValue = CStr(xlsWs.Cells(i, 2).Value)  ' B列の元の値
        origFValue = CStr(xlsWs.Cells(i, 6).Value)  ' F列の元の値
        origGValue = CStr(xlsWs.Cells(i, 7).Value)  ' G列の元の値
        
        ' 条件値が存在する場合は適用
        If hasConditions Then
            ' A列: 4桁数字を使用（存在する場合）
            If fourDigits <> "" Then
                aValue = fourDigits
            Else
                aValue = origAValue
            End If
            
            ' B列: 1桁数字を使用（存在する場合）、なければB列の値を変換
            If oneDigit <> "" Then
                bValue = oneDigit
            Else
                ' B列の変換 - 参照シートを使用して変換
                bValue = ConvertValueIfMatch(aValue, origBValue, 2, 50)
            End If
            
            ' F列: "データ"の値を使用（存在する場合）
            If condition1Value <> "" Then
                ' "データ"があれば、それをF列の値として使用
                fValue = condition1Value
            ElseIf headerInfo.Exists(origFValue) Then
                ' ヘッダー情報から変換
                fValue = headerInfo(origFValue)
            Else
                ' A列の値に基づいてF列の値を変換
                fValue = ConvertValueIfMatch(aValue, origFValue, 2, 50)
                
                ' 変換できなかった場合はデフォルト値
                If fValue = origFValue Then
                    fValue = "0000XXX"  ' デフォルト値
                End If
            End If
            
            ' G列: "処理"の値を使用（存在する場合）
            If condition2Value <> "" Then
                gValue = condition2Value
            Else
                ' A列の値に基づいてG列の値を変換
                gValue = ConvertValueIfMatch(aValue, origGValue, 2, 50)
            End If
        Else
            ' 条件がない場合は通常通り処理
            aValue = origAValue
            
            ' B列の変換 - 参照シートを使用して変換
            bValue = ConvertValueIfMatch(aValue, origBValue, 2, 50)
            
            ' F列の変換 - ヘッダー情報または参照シートを使用
            If headerInfo.Exists(origFValue) Then
                fValue = headerInfo(origFValue)
            Else
                fValue = ConvertValueIfMatch(aValue, origFValue, 2, 50)
                
                If fValue = origFValue Then
                    fValue = "0000XXX"  ' デフォルト値
                End If
            End If
            
            ' G列の値 - A列の値に基づいて変換
            gValue = ConvertValueIfMatch(aValue, origGValue, 2, 50)
        End If
        
        ' L, M列の値を取得
        lValue = CStr(xlsWs.Cells(i, 12).Value)  ' L列
        mValue = CStr(xlsWs.Cells(i, 13).Value)  ' M列
        
        ' 接頭辞の取得（XLSXのH列とI列から取得）
        Dim prefix1Code As String
        Dim prefix2Code As String
        
        prefix1Code = CStr(xlsWs.Cells(i, 8).Value)  ' H列から接頭辞1コード取得
        prefix2Code = CStr(xlsWs.Cells(i, 9).Value)  ' I列から接頭辞2コード取得
        
        ' コードから実際の値に変換（参照シートを使用）
        prefixOne = GetMappedValue(prefix1Code, "A", "B", 2, 10)  ' 接頭辞1の変換
        prefixTwo = GetMappedValue(prefix2Code, "C", "D", 2, 10)  ' 接頭辞2の変換
        
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
    
    ' カスタムデータを取得
    Dim customData1 As String
    Dim customData2 As String
    
    customData1 = ws.Range(CELL_CUSTOM_DATA1).Value
    customData2 = ws.Range(CELL_CUSTOM_DATA2).Value
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "処理中..."
    
    ' 処理実行
    ProcessFiles g_csvFilePath, g_xlsxFilePath, g_resultFilePath, customData1, customData2
    
    ' ステータス更新
    ws.Range(CELL_STATUS).Value = "完了"
End Sub