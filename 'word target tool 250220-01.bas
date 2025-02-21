Option Explicit

'// 単語から語幹を取得する関数
Private Function GetStem(word As String) As String
    Dim stem As String
    stem = LCase(Trim(word))

    '// スペースを含む場合（イディオム）はそのまま返す
    If InStr(stem, " ") > 0 Then
        GetStem = stem
        Exit Function
    End If
    
    '// 3文字以下の単語はそのままを語幹として扱う
    If Len(stem) <= 3 Then
        GetStem = stem
        Exit Function
    End If
    
    '// 名詞の接尾辞
    If stem Like "*tion": stem = Left(stem, Len(stem) - 4)
    If stem Like "*sion": stem = Left(stem, Len(stem) - 4)
    If stem Like "*ment": stem = Left(stem, Len(stem) - 4)
    If stem Like "*ity": stem = Left(stem, Len(stem) - 3)
    If stem Like "*ism": stem = Left(stem, Len(stem) - 3)
    
    '// 動詞の接尾辞
    If stem Like "*icate": stem = Left(stem, Len(stem) - 5)
    If stem Like "*ative": stem = Left(stem, Len(stem) - 5)
    If stem Like "*alize": stem = Left(stem, Len(stem) - 5)
    If stem Like "*ing": stem = Left(stem, Len(stem) - 3)
    If stem Like "*ed": stem = Left(stem, Len(stem) - 2)
    
    '// 形容詞・副詞の接尾辞
    If stem Like "*ful": stem = Left(stem, Len(stem) - 3)
    If stem Like "*ness": stem = Left(stem, Len(stem) - 4)
    If stem Like "*ly": stem = Left(stem, Len(stem) - 2)
    If stem Like "*ic": stem = Left(stem, Len(stem) - 2)
    If stem Like "*al": stem = Left(stem, Len(stem) - 2)
    
    GetStem = stem
End Function

'// 二つの単語の類似度を計算する関数
Private Function CalculateSimilarity(word1 As String, word2 As String) As Double
    '// 空文字列チェック
    If Len(word1) = 0 Then
        If Len(word2) = 0 Then
            CalculateSimilarity = 1
        Else
            CalculateSimilarity = 0
        End If
        Exit Function
    ElseIf Len(word2) = 0 Then
        CalculateSimilarity = 0
        Exit Function
    End If
    
    '// 同じ文字列の場合
    If word1 = word2 Then
        CalculateSimilarity = 1
        Exit Function
    End If
    
    '// 文字列の長さを取得
    Dim len1 As Integer, len2 As Integer
    len1 = Len(word1)
    len2 = Len(word2)
    
    '// 距離計算用の配列（必要な分だけ確保）
    Dim v0() As Integer, v1() As Integer
    ReDim v0(len2)
    ReDim v1(len2)
    
    '// 初期化
    Dim i As Integer, j As Integer
    For i = 0 To len2
        v0(i) = i
    Next i
    
    '// 距離を計算
    For i = 0 To len1 - 1
        v1(0) = i + 1
        
        For j = 0 To len2 - 1
            Dim cost As Integer
            If Mid(word1, i + 1, 1) = Mid(word2, j + 1, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            v1(j + 1) = WorksheetFunction.Min( _
                v1(j) + 1, _
                v0(j + 1) + 1, _
                v0(j) + cost)
        Next j
        
        '// v1をv0にコピー
        For j = 0 To len2
            v0(j) = v1(j)
        Next j
    Next i
    
    '// 類似度を計算（0～1の範囲）
    Dim maxLen As Integer
    maxLen = WorksheetFunction.Max(len1, len2)
    
    If maxLen = 0 Then
        CalculateSimilarity = 0
    Else
        CalculateSimilarity = 1 - (v1(len2) / maxLen)
    End If
End Function

Private Function HasSameStem(word1 As String, word2 As String) As Boolean
    If Len(Trim(word1)) = 0 Or Len(Trim(word2)) = 0 Then
        HasSameStem = False
        Exit Function
    End If
    
    '// イディオムと単語の比較の場合
    Dim hasSpace1 As Boolean, hasSpace2 As Boolean
    hasSpace1 = (InStr(word1, " ") > 0)
    hasSpace2 = (InStr(word2, " ") > 0)
    
    '// 両方ともイディオムの場合は完全一致で比較
    If hasSpace1 And hasSpace2 Then
        HasSameStem = (LCase(Trim(word1)) = LCase(Trim(word2)))
        Exit Function
    End If
    
    '// イディオムと単語の比較の場合、各単語を分解して比較
    If hasSpace1 Or hasSpace2 Then
        Dim idiom As String, singleWord As String
        Dim idiomWords() As String
        
        If hasSpace1 Then
            idiomWords = Split(LCase(Trim(word1)))
            singleWord = LCase(Trim(word2))
        Else
            idiomWords = Split(LCase(Trim(word2)))
            singleWord = LCase(Trim(word1))
        End If
        
        '// イディオム内の各単語と比較
        Dim i As Long
        For i = 0 To UBound(idiomWords)
            '// 3文字以下の単語は完全一致のみ
            If Len(idiomWords(i)) <= 3 Then
                If idiomWords(i) = singleWord Then
                    HasSameStem = True
                    Exit Function
                End If
            Else
                Dim stem1 As String, stem2 As String
                stem1 = GetStem(idiomWords(i))
                stem2 = GetStem(singleWord)
                
                If CalculateSimilarity(stem1, stem2) >= 0.8 Then
                    HasSameStem = True
                    Exit Function
                End If
            End If
        Next i
        
        HasSameStem = False
        Exit Function
    End If
    
    '// 単語同士の比較（既存のロジック）
    Dim stem1Final As String, stem2Final As String
    stem1Final = GetStem(word1)
    stem2Final = GetStem(word2)
    
    If Len(stem1Final) <= 3 Or Len(stem2Final) <= 3 Then
        HasSameStem = (stem1Final = stem2Final)
    Else
        HasSameStem = (CalculateSimilarity(stem1Final, stem2Final) >= 0.8)
    End If
End Function




Public Sub ProcessWords()
    On Error GoTo ErrorHandler
    
    Debug.Print "ProcessWords 開始: " & Now()
    Dim startTime As Double
    startTime = Timer
    
    '// シートの設定
    Dim wsA As Worksheet, wsB As Worksheet, wsC As Worksheet
    Set wsA = ThisWorkbook.Sheets("単語リスト")
    Set wsB = ThisWorkbook.Sheets("ターゲット候補")
    Set wsC = ThisWorkbook.Sheets("処理ログ")
    
    '// アプリケーション設定を保存
    Dim oldStatusBar As Boolean
    Dim oldScreenUpdating As Boolean
    oldStatusBar = Application.DisplayStatusBar
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    
    '// シートCをクリア
    wsC.Cells.Clear
    Debug.Print "シートCをクリア完了"
    
    '// ヘッダー行の設定
    Dim startRow As Long
    startRow = 2  '// ヘッダー行を除外
    
    '// データ範囲の取得
    Dim lastRowA As Long, lastRowB As Long
    lastRowA = wsA.Cells(wsA.Rows.Count, "D").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    
    Debug.Print "データ範囲取得完了 - シートA最終行: " & lastRowA & ", シートB最終行: " & lastRowB
    
    '// 入力チェック
    If lastRowB < startRow Then
        Debug.Print "エラー: シートBにデータが存在しません"
        MsgBox "シートBにデータが存在しません。", vbExclamation
        GoTo Cleanup
    End If
    
    If lastRowA < startRow Then
        Debug.Print "エラー: シートAにデータが存在しません"
        MsgBox "シートAにデータが存在しません。", vbExclamation
        GoTo Cleanup
    End If
    
    '// ヘッダーの設定
    wsC.Cells(1, "A").Value = "対象単語"
    wsC.Cells(1, "B").Value = "語幹"
    wsC.Cells(1, "C").Value = "候補単語"
    wsC.Cells(1, "D").Value = "候補語幹"
    wsC.Cells(1, "E").Value = "最終結果"
    Debug.Print "ヘッダー設定完了"
    
    '// ステップ1: シートBのA列をシートCのA列にコピーし、B列に語幹を取得
    Application.StatusBar = "ステップ1: コピー処理を開始します..."
    Debug.Print "ステップ1開始: A列コピー"
    wsB.Range("A" & startRow & ":A" & lastRowB).Copy wsC.Range("A" & startRow)
    
    Dim i As Long, j As Long
    For i = startRow To lastRowB
        Application.StatusBar = "ステップ1: 語幹取得中 " & Format((i - startRow + 1) / (lastRowB - startRow + 1), "0%") & " 完了..."
        wsC.Cells(i, "B").Value = GetStem(wsC.Cells(i, "A").Value)
        If i Mod 100 = 0 Then Debug.Print "ステップ1進捗: " & i & "/" & lastRowB & "完了"
    Next i
    Debug.Print "ステップ1完了"
    
    '// ステップ2: シートAのD列と比較して条件を満たす単語をC列に保存
    Application.StatusBar = "ステップ2: 単語比較を開始します..."
    Debug.Print "ステップ2開始: 単語比較"
    
    Dim resultCount As Long
    resultCount = startRow - 1  '// resultCountの初期値をヘッダー行の次に設定
    
    For i = startRow To lastRowA
        Application.StatusBar = "ステップ2: 単語比較中 " & Format((i - startRow + 1) / (lastRowA - startRow + 1), "0%") & " 完了..."
        
        Dim isValid As Boolean
        isValid = True
        
        For j = startRow To lastRowB
            If HasSameStem(wsA.Cells(i, "D").Value, wsC.Cells(j, "A").Value) Then
                isValid = False
                Exit For
            End If
        Next j
        
        If isValid Then
            resultCount = resultCount + 1
            wsC.Cells(resultCount, "C").Value = wsA.Cells(i, "D").Value
        End If
        
        If i Mod 100 = 0 Then Debug.Print "ステップ2進捗: " & i & "/" & lastRowA & "完了, 現在の結果数: " & (resultCount - startRow + 1)
    Next i
    Debug.Print "ステップ2完了. 総結果数: " & (resultCount - startRow + 1)
    
    '// ステップ3: C列の単語の語幹をD列に保存
    If resultCount >= startRow Then
        Application.StatusBar = "ステップ3: 結果単語の語幹取得を開始します..."
        Debug.Print "ステップ3開始: 結果単語の語幹取得"
        
        For i = startRow To resultCount
            Application.StatusBar = "ステップ3: 語幹取得中 " & Format((i - startRow + 1) / (resultCount - startRow + 1), "0%") & " 完了..."
            If wsC.Cells(i, "C").Value <> "" Then
                wsC.Cells(i, "D").Value = GetStem(wsC.Cells(i, "C").Value)
            End If
            If i Mod 100 = 0 Then Debug.Print "ステップ3進捗: " & i & "/" & resultCount & "完了"
        Next i
        Debug.Print "ステップ3完了"
        
        '// ステップ4: C,D列の単語を比較して同じ語幹の単語から最短のものをE列に保存
        Application.StatusBar = "ステップ4: 最終結果作成を開始します..."
        Debug.Print "ステップ4開始: 最終結果作成"
        
        wsC.Range("C" & startRow & ":C" & resultCount).Copy wsC.Range("E" & startRow)
        
        For i = startRow To resultCount
            Application.StatusBar = "ステップ4: 最終処理中 " & Format((i - startRow + 1) / (resultCount - startRow + 1), "0%") & " 完了..."
            If wsC.Cells(i, "E").Value <> "" Then
                Dim shortestWord As String
                shortestWord = wsC.Cells(i, "E").Value
                Dim shortestIndex As Long
                shortestIndex = i
                
                For j = startRow To resultCount
                    If j <> i And wsC.Cells(j, "E").Value <> "" Then
                        If HasSameStem(wsC.Cells(i, "E").Value, wsC.Cells(j, "E").Value) Then
                            If Len(wsC.Cells(j, "E").Value) < Len(shortestWord) Then
                                shortestWord = wsC.Cells(j, "E").Value
                                shortestIndex = j
                            End If
                        End If
                    End If
                Next j
                
                If shortestIndex <> i Then
                    wsC.Cells(i, "E").ClearContents
                Else
                    For j = startRow To resultCount
                        If j <> i And wsC.Cells(j, "E").Value <> "" Then
                            If HasSameStem(shortestWord, wsC.Cells(j, "E").Value) Then
                                wsC.Cells(j, "E").ClearContents
                            End If
                        End If
                    Next j
                End If
            End If
            If i Mod 100 = 0 Then Debug.Print "ステップ4進捗: " & i & "/" & resultCount & "完了"
        Next i
        Debug.Print "ステップ4完了"
    Else
        Debug.Print "警告: 処理可能な単語が見つかりません"
        MsgBox "処理可能な単語が見つかりませんでした。", vbInformation
    End If
    
    Debug.Print "処理時間: " & Format(Timer - startTime, "0.00") & "秒"
    GoTo Cleanup
    
ErrorHandler:
    Debug.Print "エラー発生: " & Now()
    Debug.Print "エラー番号: " & Err.Number
    Debug.Print "エラーの説明: " & Err.Description
    Debug.Print "エラー発生箇所: " & Erl
    
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, _
           vbCritical + vbOKOnly, _
           "エラー"
    
Cleanup:
    '// アプリケーション設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayStatusBar = oldStatusBar
    
    Debug.Print "ProcessWords 終了: " & Now()
End Sub