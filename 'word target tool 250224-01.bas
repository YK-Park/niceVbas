'word target tool 250224-01.bas
Option Explicit

'// 語幹を取得する関数
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
    
    '// レーベンシュタイン距離の計算
    Dim len1 As Integer, len2 As Integer
    len1 = Len(word1)
    len2 = Len(word2)
    
    Dim v0() As Integer, v1() As Integer
    ReDim v0(len2)
    ReDim v1(len2)
    
    Dim i As Integer, j As Integer
    For i = 0 To len2
        v0(i) = i
    Next i
    
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
        
        For j = 0 To len2
            v0(j) = v1(j)
        Next j
    Next i
    
    Dim maxLen As Integer
    maxLen = WorksheetFunction.Max(len1, len2)
    
    If maxLen = 0 Then
        CalculateSimilarity = 0
    Else
        CalculateSimilarity = 1 - (v1(len2) / maxLen)
    End If
End Function

'// 同じ語幹を持つかチェックする関数
Private Function HasSameStem(word1 As String, word2 As String, similarityThreshold As Double) As Boolean
    If Len(Trim(word1)) = 0 Or Len(Trim(word2)) = 0 Then
        HasSameStem = False
        Exit Function
    End If
    
    '// イディオムと単語の比較
    Dim hasSpace1 As Boolean, hasSpace2 As Boolean
    hasSpace1 = (InStr(word1, " ") > 0)
    hasSpace2 = (InStr(word2, " ") > 0)
    
    '// 両方イディオムの場合
    If hasSpace1 And hasSpace2 Then
        HasSameStem = (LCase(Trim(word1)) = LCase(Trim(word2)))
        Exit Function
    End If
    
    '// イディオムと単語の比較
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
        
        Dim i As Long
        For i = 0 To UBound(idiomWords)
            If Len(idiomWords(i)) <= 3 Then
                If idiomWords(i) = singleWord Then
                    HasSameStem = True
                    Exit Function
                End If
            Else
                Dim stem1 As String, stem2 As String
                stem1 = GetStem(idiomWords(i))
                stem2 = GetStem(singleWord)
                
                If CheckStemInclusion(stem1, stem2) Then
                    HasSameStem = True
                    Exit Function
                End If
                
                If CalculateSimilarity(stem1, stem2) >= similarityThreshold Then
                    HasSameStem = True
                    Exit Function
                End If
            End If
        Next i
        
        HasSameStem = False
        Exit Function
    End If
    
    '// 単語同士の比較
    Dim stem1Final As String, stem2Final As String
    stem1Final = GetStem(word1)
    stem2Final = GetStem(word2)
    
    If Len(stem1Final) <= 3 Or Len(stem2Final) <= 3 Then
        If Len(stem1Final) < Len(stem2Final) Then
            HasSameStem = InStr(stem2Final, stem1Final) > 0
        Else
            HasSameStem = InStr(stem1Final, stem2Final) > 0
        End If
    Else
        If CheckStemInclusion(stem1Final, stem2Final) Then
            HasSameStem = True
            Exit Function
        End If
        
        HasSameStem = (CalculateSimilarity(stem1Final, stem2Final) >= similarityThreshold)
    End If
End Function

Private Function CheckStemInclusion(stem1 As String, stem2 As String) As Boolean
    Dim shortStem As String, longStem As String
    
    If Len(stem1) < Len(stem2) Then
        shortStem = stem1
        longStem = stem2
    Else
        shortStem = stem2
        longStem = stem1
    End If
    
    CheckStemInclusion = (InStr(1, longStem, shortStem) > 0)
End Function

Public Sub ProcessWords()
    On Error GoTo ErrorHandler
    
    Debug.Print "ProcessWords 開始: " & Now()
    Dim startTime As Double
    startTime = Timer
    
    '// シートの設定
    Dim wsA As Worksheet, wsB As Worksheet
    Set wsA = ThisWorkbook.Sheets("単語リスト")
    Set wsB = ThisWorkbook.Sheets("ターゲット候補")
    
    '// アプリケーション設定を保存
    Dim oldStatusBar As Boolean
    Dim oldScreenUpdating As Boolean
    oldStatusBar = Application.DisplayStatusBar
    oldScreenUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    
    '// 類似度閾値の取得
    Dim similarityThreshold As Double
    similarityThreshold = wsB.Range("B1").Value
    If similarityThreshold <= 0 Or similarityThreshold > 1 Then
        similarityThreshold = 0.8 '// デフォルト値
    End If
    
    '// ヘッダー行の設定
    Dim startRow As Long
    startRow = 3  '// ヘッダー行とボタン用の行を除外
    
    '// データ範囲の取得
    Dim lastRowA As Long, lastRowB As Long
    lastRowA = wsA.Cells(wsA.Rows.Count, "D").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "A").End(xlUp).Row
    
    '// 入力チェック
    If lastRowB < startRow Then
        MsgBox "シートBにデータが存在しません。", vbExclamation
        GoTo Cleanup
    End If
    
    If lastRowA < startRow Then
        MsgBox "シートAにデータが存在しません。", vbExclamation
        GoTo Cleanup
    End If
    
    '// 結果格納用の型定義
    Type WordResult
        LevelNum As String    '// 級番号
        UniqueNum As String   '// ユニーク番号
        Level As String       '// 級
        Word As String        '// ターゲット単語
        PartOfSpeech As String '// 品詞
        Category As String    '// 出題区分
        IsValid As Boolean    '// 有効フラグ
    End Type
    
    '// 配列の準備
    Dim results() As WordResult
    ReDim results(1 To lastRowA - startRow + 1)
    Dim resultCount As Long
    resultCount = 0
    
    '// シートAの単語を処理して配列に保存
    Dim i As Long, j As Long
    For i = startRow To lastRowA
        Application.StatusBar = "単語処理中: " & Format((i - startRow + 1) / (lastRowA - startRow + 1), "0%")
        
        Dim currentWord As String
        currentWord = wsA.Cells(i, "D").Value
        
        '// シートBのA列の単語とチェック
        Dim isValid As Boolean
        isValid = True
        
        For j = startRow To lastRowB
            If wsB.Cells(j, "A").Value <> "" Then
                If HasSameStem(currentWord, wsB.Cells(j, "A").Value, similarityThreshold) Then
                    isValid = False
                    Exit For
                End If
            End If
        Next j
        
        '// すでに配列にある単語とチェック
        If isValid Then
            For j = 1 To resultCount
                If results(j).IsValid Then
                    If HasSameStem(currentWord, results(j).Word, similarityThreshold) Then
                        isValid = False
                        Exit For
                    End If
                End If
            Next j
        End If
        
        '// 条件を満たす場合、配列に追加
        If isValid Then
            resultCount = resultCount + 1
            With results(resultCount)
                .LevelNum = wsA.Cells(i, "A").Value
                .UniqueNum = wsA.Cells(i, "B").Value
                .Level = wsA.Cells(i, "C").Value
                .Word = wsA.Cells(i, "D").Value
                .PartOfSpeech = wsA.Cells(i, "E").Value
                .Category = wsA.Cells(i, "F").Value
                .IsValid = True
            End With
        End If
    Next i
    
    '// 同じ語幹を持つ単語のうち、最短のものだけを残す
    If resultCount > 0 Then
        For i = 1 To resultCount
            If results(i).IsValid Then
                Dim shortestWord As String
                shortestWord = results(i).Word
                Dim shortestIndex As Long
                shortestIndex = i
                
                For j = i + 1 To resultCount
                    If results(j).IsValid Then
                        If HasSameStem(shortestWord, results(j).Word, similarityThreshold) Then
                            If Len(results(j).Word) < Len(shortestWord) Then
                                shortestWord = results(j).Word
                                shortestIndex = j
                            End If
                        End If
                    End If
                Next j
                
                If shortestIndex <> i Then
                    results(i).IsValid = False
                Else
                    For j = i + 1 To resultCount
                        If results(j).IsValid Then
                            If HasSameStem(shortestWord, results(j).Word, similarityThreshold) Then
                                results(j).IsValid = False
                            End If
                        End If
                    Next j
                End If
            End If
        Next i
        
        '// 結果をワークシートに書き込む
        Application.StatusBar = "結果を書き込み中..."
        wsB.Range("C" & startRow & ":H" & wsB.Rows.Count).ClearContents
        
        Dim writeRow As Long
        writeRow = startRow
        
        For i = 1 To resultCount
            If results(i).IsValid Then
                With wsB
                    .Cells(writeRow, "C").Value = results(i).LevelNum
                    .Cells(writeRow, "D").Value = results(i).UniqueNum
                    .Cells(writeRow, "E").Value = results(i).Level
                    .Cells(writeRow, "F").Value = results(i).Word
                    .Cells(writeRow, "G").Value = results(i).PartOfSpeech
                    .Cells(writeRow, "H").Value = results(i).Category
                End With
                writeRow = writeRow + 1
            End If
        Next i
    End If
    
    Debug.Print "処理時間: " & Format(Timer - startTime, "0.00") & "秒"
    MsgBox "処理が完了しました。", vbInformation
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