'word target tool 250211-01
Option Explicit

'*** 類似度を計算するための定数 ***
Private Const SIMILARITY_THRESHOLD As Double = 0.7    '類似度の閾値
Private Const MIN_LENGTH As Integer = 4               '比較する最小文字数
Private Const PHRASE_SIMILARITY_THRESHOLD As Double = 0.8  'フレーズの類似度閾値

'*** 語幹の抽出用の配列定数 ***
Private Const COMMON_SUFFIXES As String = "ed,ing,s,es,er,est,ly,ment,ness,ful,less,able,ible,al,ial,ic,ical,ish,like,ive,ative,itive"
Private Const BE_FORMS As String = "be,am,is,are,was,were,been,being"

'*** 単語検索と除外処理を行うメインの関数 ***
Public Sub SearchRelatedWords()
    '*** 変数の宣言 ***
    Dim wsTarget As Worksheet    '検索対象のワークシート
    Dim wsList As Worksheet      '単語リストのワークシート
    Dim targetPhrase As String   '検索する単語またはフレーズ
    Dim lastRow As Long          'データの最終行
    Dim i As Long                'ループカウンター
    Dim resultRow As Long        '結果を書き込む行番号
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    '*** ワークシートの設定 ***
    Set wsTarget = ThisWorkbook.Sheets("ターゲット候補")
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** 検索する単語/フレーズの取得と前処理 ***
    targetPhrase = LCase(Trim(wsTarget.Range("D1").Value))
    If targetPhrase = "" Then
        MsgBox "検索する単語を入力してください。", vbExclamation
        Exit Sub
    End If
    
    '*** 結果エリアのクリア ***
    ClearResults wsTarget
    
    '*** データの最終行を取得 ***
    lastRow = wsList.Cells(wsList.Rows.Count, "D").End(xlUp).Row
    
    '*** 結果の書き込み開始行 ***
    resultRow = 1
    
    '*** 単語リストの検索処理 ***
    For i = 2 To lastRow
        Dim currentPhrase As String
        currentPhrase = LCase(Trim(wsList.Cells(i, "D").Value))
        
        '*** 検索条件に合致する単語/フレーズの処理 ***
        If currentPhrase <> "" Then
            If Not IsExcludedPhrase(currentPhrase, targetPhrase) Then
                '*** 結果の書き込み ***
                WriteResult wsList, wsTarget, i, resultRow
                resultRow = resultRow + 1
            End If
        End If
    Next i
    
    '*** 結果の書式設定 ***
    If resultRow > 1 Then
        FormatResults wsTarget, resultRow - 1
    End If
    
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

'*** 除外すべきフレーズかどうかを判定する関数 ***
'*** フレーズ内の各単語を取得 ***
Private Function SplitPhrase(ByVal phrase As String) As String()
    If Trim(phrase) = "" Then
        ReDim result(0) As String
        result(0) = ""
        SplitPhrase = result
        Exit Function
    End If
    SplitPhrase = Split(LCase(Trim(phrase)), " ")
End Function

'*** 除外すべきフレーズかどうかを判定する関数 ***
Private Function IsExcludedPhrase(ByVal phrase As String, ByVal target As String) As Boolean
    '*** 空文字チェック ***
    If Trim(phrase) = "" Or Trim(target) = "" Then
        IsExcludedPhrase = False
        Exit Function
    End If
    
    '*** 完全一致の場合は除外 ***
    If phrase = target Then
        IsExcludedPhrase = True
        Exit Function
    End If
    
    '*** フレーズの場合の処理 ***
    If IsPhrase(phrase) Then
        '*** フレーズ内の各単語を取得 ***
        Dim phraseWords() As String
        phraseWords = SplitPhrase(phrase)
        
        If UBound(phraseWords) < 0 Then
            IsExcludedPhrase = False
            Exit Function
        End If
       
        '*** be動詞を除いた実質的な単語を抽出 ***
        Dim i As Long
        For i = 0 To UBound(phraseWords)
            If phraseWords(i) <> "" Then  ' 空文字チェックを追加
                '*** be動詞はスキップ ***
                If Not IsBeForm(phraseWords(i)) Then
                    '*** 語幹を抽出して比較 ***
                    Dim stemWord As String
                    stemWord = GetWordStem(phraseWords(i))
                    
                    '*** 検索語の語幹を抽出 ***
                    Dim targetStem As String
                    targetStem = GetWordStem(target)
                    
                    '*** 語幹が一致または高い類似度の場合は除外 ***
                    If stemWord <> "" And targetStem <> "" Then  ' 空文字チェックを追加
                        If stemWord = targetStem Or _
                           (Len(stemWord) >= MIN_LENGTH And Len(targetStem) >= MIN_LENGTH And _
                            CalculateSimilarity(stemWord, targetStem) >= SIMILARITY_THRESHOLD) Then
                            IsExcludedPhrase = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i
        
        IsExcludedPhrase = False
    Else
        '*** 単語の場合は通常の類似度チェック ***
        If Len(phrase) >= MIN_LENGTH And Len(target) >= MIN_LENGTH Then
            Dim similarity As Double
            similarity = CalculateSimilarity(GetWordStem(phrase), GetWordStem(target))
            IsExcludedPhrase = (similarity >= SIMILARITY_THRESHOLD)
        Else
            IsExcludedPhrase = False
        End If
    End If
End Function

'*** フレーズかどうかを判定する関数 ***
Private Function IsPhrase(ByVal text As String) As Boolean
    IsPhrase = (InStr(text, " ") > 0)
End Function

'*** 語幹を抽出する関数 ***
Private Function GetWordStem(ByVal word As String) As String
    Dim stem As String
    Dim suffixes() As String
    
    stem = LCase(Trim(word))
    suffixes = Split(COMMON_SUFFIXES, ",")
    
    Dim i As Long
    For i = 0 To UBound(suffixes)
        If Len(stem) > Len(suffixes(i)) + 2 Then  '最低3文字は残す
            If Right(stem, Len(suffixes(i))) = suffixes(i) Then
                stem = Left(stem, Len(stem) - Len(suffixes(i)))
                Exit For
            End If
        End If
    Next i
    
    GetWordStem = stem
End Function

'*** be動詞かどうかを判定する関数 ***
Private Function IsBeForm(ByVal word As String) As Boolean
    Dim beForms() As String
    beForms = Split(BE_FORMS, ",")
    
    Dim i As Long
    For i = 0 To UBound(beForms)
        If LCase(word) = beForms(i) Then
            IsBeForm = True
            Exit Function
        End If
    Next i
    
    IsBeForm = False
End Function

'*** 類似度を計算する関数 ***
Private Function CalculateSimilarity(ByVal word1 As String, ByVal word2 As String) As Double
    Dim distance As Long
    distance = LevenshteinDistance(word1, word2)
    CalculateSimilarity = 1 - (distance / Application.Max(Len(word1), Len(word2)))
End Function

'*** レーベンシュタイン距離を計算する関数 ***
Private Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    Dim i As Long, j As Long
    Dim len1 As Long, len2 As Long
    Dim matrix() As Long
    Dim cost As Long
    
    len1 = Len(s1)
    len2 = Len(s2)
    
    ReDim matrix(len1, len2)
    
    '*** 行列の初期化 ***
    For i = 0 To len1
        matrix(i, 0) = i
    Next i
    
    For j = 0 To len2
        matrix(0, j) = j
    Next j
    
    '*** 距離の計算 ***
    For i = 1 To len1
        For j = 1 To len2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            matrix(i, j) = Application.Min( _
                matrix(i - 1, j) + 1, _
                matrix(i, j - 1) + 1, _
                matrix(i - 1, j - 1) + cost)
        Next j
    Next i
    
    LevenshteinDistance = matrix(len1, len2)
End Function

'*** 結果をターゲットシートに書き込む関数 ***
Private Sub WriteResult(ByVal sourceWs As Worksheet, ByVal targetWs As Worksheet, _
                       ByVal sourceRow As Long, ByVal resultRow As Long)
    targetWs.Cells(resultRow, "J").Value = sourceWs.Cells(sourceRow, "A").Value '級番号
    targetWs.Cells(resultRow, "K").Value = sourceWs.Cells(sourceRow, "B").Value 'ユニーク番号
    targetWs.Cells(resultRow, "L").Value = sourceWs.Cells(sourceRow, "C").Value '級
    targetWs.Cells(resultRow, "M").Value = sourceWs.Cells(sourceRow, "D").Value 'ターゲット単語
    targetWs.Cells(resultRow, "N").Value = sourceWs.Cells(sourceRow, "E").Value '品詞
    targetWs.Cells(resultRow, "O").Value = sourceWs.Cells(sourceRow, "F").Value '出題区分
End Sub

'*** 結果エリアをクリアする関数 ***
Private Sub ClearResults(ByVal ws As Worksheet)
    ws.Range("J:O").ClearContents
End Sub

'*** 結果の書式を設定する関数 ***
Private Sub FormatResults(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range(ws.Cells(1, "J"), ws.Cells(lastRow, "O"))
        .Borders.LineStyle = xlContinuous
        .Font.Name = "メイリオ"
        .Font.Size = 11
        .EntireColumn.AutoFit
    End With
End Sub