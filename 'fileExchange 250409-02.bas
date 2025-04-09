'fileExchange 250409-02
' 大文字の単語、熟語、派生語を置換し、セルに色を付ける関数
Public Sub ReplaceUppercaseWords()
    ' 定数定義
    Const SHEET_NAME As String = "CSV"
    Const START_ROW As Long = 2
    Const SEARCH_COL As String = "A"
    Const TARGET_COL As String = "E"
    Const PLACEHOLDER As String = "(          )"
    
    ' 色の定義
    Const COLOR_DERIVED_WORDS As Long = RGB(255, 255, 0)    ' 黄色 - 派生語が見つかった場合（例：ADDsなど）
    Const COLOR_EXACT_MATCH As Long = RGB(173, 216, 230)    ' 水色 - 完全一致のみの場合
    Const COLOR_NO_MATCH As Long = RGB(255, 200, 200)       ' 薄い赤 - 置換がなかった場合
    Const COLOR_SIMILAR_MATCH As Long = RGB(144, 238, 144)  ' 薄い緑 - 類似単語や熟語が見つかった場合
    
    ' 変数宣言
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim searchWord As String
    Dim targetText As String
    Dim regExExact As Object   ' 完全一致用
    Dim regExDerived As Object ' 派生語用
    Dim statusBarText As String
    Dim progressPct As Double
    Dim changedCount As Long
    Dim exactMatchCount As Long
    Dim derivedMatchCount As Long
    Dim noMatchCount As Long
    Dim isPhrase As Boolean    ' 熟語判定フラグ
    
    ' 改行を処理するための変数
    Dim textLines As Variant
    Dim lineIdx As Long
    Dim processedText As String
    Dim hasExactMatch As Boolean
    Dim hasDerivedMatch As Boolean
    Dim wasReplaced As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 元のステータスバーテキストを保存
    statusBarText = Application.StatusBar
    
    ' パフォーマンス最適化の設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' CSVシートの確認
    If Not SheetExists(SHEET_NAME) Then
        MsgBox "シート「" & SHEET_NAME & "」が見つかりません。", vbExclamation
        GoTo CleanExit
    End If
    
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    ' データの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, SEARCH_COL).End(xlUp).Row
    
    ' データが存在しない場合
    If lastRow < START_ROW Then
        MsgBox "処理するデータがありません。", vbInformation
        GoTo CleanExit
    End If
    
    ' 正規表現オブジェクトを作成（2つ用意）
    Set regExExact = CreateObject("VBScript.RegExp")
    regExExact.Global = True
    
    Set regExDerived = CreateObject("VBScript.RegExp")
    regExDerived.Global = True
    
    ' 進捗表示の準備
    Application.StatusBar = "大文字の単語を置換しています...0%"
    changedCount = 0
    exactMatchCount = 0
    derivedMatchCount = 0
    noMatchCount = 0
    
    ' 各行の処理
    For i = START_ROW To lastRow
        ' 進捗状況を表示
        If i Mod 10 = 0 Or i = lastRow Then
            progressPct = CDbl(i - START_ROW + 1) / CDbl(lastRow - START_ROW + 1) * 100
            Application.StatusBar = "大文字の単語を置換しています..." & Format(progressPct, "0") & "%"
        End If
        
        ' データの取得
        searchWord = ws.Range(SEARCH_COL & i).Value
        targetText = ws.Range(TARGET_COL & i).Value
        
        ' セルの背景色をリセット
        ws.Range(TARGET_COL & i).Interior.ColorIndex = xlNone
        
        ' フラグの初期化
        hasExactMatch = False
        hasDerivedMatch = False
        wasReplaced = False
        
        ' 検索単語と対象テキストが存在し、検索単語が大文字の場合
        If searchWord <> "" And targetText <> "" And searchWord = UCase(searchWord) Then
            ' 単語か熟語かの判定 (スペースが含まれていれば熟語)
            isPhrase = InStr(searchWord, " ") > 0
            
            ' 特殊文字のエスケープは不要なので削除
            
            ' 正規表現のパターンを設定
            ' パターンを単語か熟語かに応じて設定
            If isPhrase Then
                ' 熟語の場合 - 完全一致
                regExExact.Pattern = searchWord
                
                ' 熟語の派生形 (例: HIGH SCHOOL -> HIGH SCHOOLS)
                ' 最後の単語の後ろに文字が続くパターン
                regExDerived.Pattern = searchWord & "[^\s]+"
            Else
                ' 単語の場合 - 完全一致 (単語境界\bを使用)
                regExExact.Pattern = "\b" & searchWord & "\b"
                
                ' 単語の派生形 (例: ADD -> ADDs, ADDED)
                ' 単語の先頭が一致し、後に文字が続くパターン（単語境界\bを使用）
                regExDerived.Pattern = "\b" & searchWord & "[^\s]*\b"
            End If
            
            ' 改行で分割して各行を処理
            textLines = Split(targetText, vbCrLf)
            processedText = ""
            
            ' 各行を個別に処理
            For lineIdx = LBound(textLines) To UBound(textLines)
                Dim originalLine As String
                Dim processedLine As String
                originalLine = textLines(lineIdx)
                processedLine = originalLine
                
                ' まず完全一致を検出・置換
                If regExExact.Test(processedLine) Then
                    ' 完全一致のマッチングと置換処理
                    If isPhrase Then
                        ' 熟語の完全一致の場合
                        processedLine = ReplacePhrase(processedLine, searchWord, PLACEHOLDER)
                    Else
                        ' 単語の完全一致の場合
                        processedLine = regExExact.Replace(processedLine, PLACEHOLDER)
                    End If
                    
                    hasExactMatch = True
                    wasReplaced = True
                
                ' 次に派生語を検出・置換 (完全一致がない場合のみ)
                ElseIf regExDerived.Test(processedLine) Then
                    ' 派生語のマッチングと置換処理
                    If isPhrase Then
                        ' 熟語の派生形の場合
                        processedLine = ReplacePhrase(processedLine, searchWord, PLACEHOLDER)
                    Else
                        ' 単語の派生形の場合
                        processedLine = regExDerived.Replace(processedLine, PLACEHOLDER)
                    End If
                    
                    hasDerivedMatch = True
                    wasReplaced = True
                End If
                
                textLines(lineIdx) = processedLine
            Next lineIdx
            
            ' 処理済みの行を結合
            For lineIdx = LBound(textLines) To UBound(textLines)
                If lineIdx = LBound(textLines) Then
                    processedText = textLines(lineIdx)
                Else
                    processedText = processedText & vbCrLf & textLines(lineIdx)
                End If
            Next lineIdx
            
            ' 処理結果をセルに書き込み
            ws.Range(TARGET_COL & i).Value = processedText
            
            ' セルの色を設定
            If wasReplaced Then
                changedCount = changedCount + 1
                
                If hasDerivedMatch Then
                    ' 派生語が見つかった場合は黄色
                    ws.Range(TARGET_COL & i).Interior.Color = COLOR_DERIVED_WORDS
                    derivedMatchCount = derivedMatchCount + 1
                ElseIf hasExactMatch Then
                    ' 完全一致のみの場合は水色
                    ws.Range(TARGET_COL & i).Interior.Color = COLOR_EXACT_MATCH
                    exactMatchCount = exactMatchCount + 1
                End If
            Else
                ' 置換がなかった場合は薄い赤
                ws.Range(TARGET_COL & i).Interior.Color = COLOR_NO_MATCH
                noMatchCount = noMatchCount + 1
                
                ' 大文字小文字を区別せずに検索して置換
                ' 一致がなかった場合でも、小文字でマッチするか試みる
                Dim lowerCaseMatch As Boolean
                lowerCaseMatch = ProcessCaseInsensitiveMatch(processedText, searchWord, PLACEHOLDER)
                
                If lowerCaseMatch Then
                    ' 小文字で一致があった場合、処理結果を更新
                    ws.Range(TARGET_COL & i).Value = processedText
                    ' 色は薄い赤のままとする
                End If
            End If
        End If
    Next i
    
    ' 完了メッセージ
    MsgBox "処理が完了しました。" & vbCrLf & _
           "合計 " & changedCount & " 箇所の単語や熟語を置換しました。" & vbCrLf & _
           "- 完全一致のみ: " & exactMatchCount & " 件（水色）" & vbCrLf & _
           "- 派生語を含む: " & derivedMatchCount & " 件（黄色 - 例：ADDsなど）" & vbCrLf & _
           "- 置換なし: " & noMatchCount & " 件（薄い赤）", vbInformation, "処理完了"

CleanExit:
    ' 設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = statusBarText
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラーコード: " & Err.Number & vbCrLf & _
           "エラー発生箇所: " & Erl, vbCritical, "エラー"
    
    GoTo CleanExit
End Sub

' 対소文字区別せず一致する単語を処理する関数
Private Function ProcessCaseInsensitiveMatch(ByRef text As String, ByVal searchWord As String, ByVal replacement As String) As Boolean
    Dim result As Boolean
    Dim regEx As Object
    Dim originalText As String
    
    ' 元のテキストを保存
    originalText = text
    
    ' 正規表現オブジェクトを作成
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True ' 大文字小文字を区別しない
    
    ' 検索語が熟語かどうか確認
    Dim isPhrase As Boolean
    isPhrase = InStr(searchWord, " ") > 0
    
    ' 正規表現パターンを設定
    If isPhrase Then
        ' 熟語の場合
        regEx.Pattern = searchWord & "[A-Za-z]*"
    Else
        ' 単語の場合、単語境界を含む
        regEx.Pattern = "\b" & searchWord & "[A-Za-z]*\b"
    End If
    
    ' パターンに一致する部分を置換
    text = regEx.Replace(text, replacement)
    
    ' 置換が行われたかどうかを確認
    result = (text <> originalText)
    
    ProcessCaseInsensitiveMatch = result
End Function

' 前後の文脈を保持しながら置換するための関数
Private Function ReplacementWithContext(ByVal match As Object) As String
    ' 単純に完全置換するため、前後の文脈は保持せずにプレースホルダーに置換
    ReplacementWithContext = "(          )"
End Function

' 熟語を置換する関数
Private Function ReplacePhrase(ByVal text As String, ByVal phrase As String, ByVal replacement As String) As String
    Dim result As String
    result = text
    
    ' 大文字に統一して比較
    Dim upperText As String
    upperText = UCase(text)
    
    ' 熟語の位置を検索
    Dim pos As Integer
    pos = InStr(upperText, phrase)
    
    ' 見つかった場合は置換
    If pos > 0 Then
        ' 派生形を考慮して、熟語の後に続く文字も含めて置換
        Dim endPos As Integer
        endPos = pos + Len(phrase)
        
        ' 熟語の後に続く文字を検出
        Do While endPos <= Len(upperText)
            Dim nextChar As String
            nextChar = Mid(upperText, endPos, 1)
            
            ' 単語の一部として認められる文字かチェック
            If nextChar Like "[A-Z0-9]" Then
                endPos = endPos + 1
            Else
                Exit Do
            End If
        Loop
        
        ' 原文から対象部分を切り出して置換
        Dim targetPhrase As String
        targetPhrase = Mid(text, pos, endPos - pos)
        result = Replace(result, targetPhrase, replacement)
    End If
    
    ReplacePhrase = result
End Function

' シートが存在するかを確認する関数
Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function