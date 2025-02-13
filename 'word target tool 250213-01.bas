'word target tool 250213-01 
'세가지 방식 동시에 사용
'*** 共通の接尾辞リスト ***
Public Const COMMON_SUFFIXES As String = "s,ed,ing,ly,er,est,ment,ness,ful,less,able,ible,al,ial,y,ify,ize,ise,ous,ious,ive,ative,itive"

'*** ステミング方式を指定する列挙型 ***
Public Enum StemmingMethod
    Porter = 1
    Levenshtein = 2
    Hybrid = 3
End Enum

'*** ステミング設定用の型 ***
Private Type StemmingConfig
    Method As StemmingMethod
    LevenshteinThreshold As Double  'レーベンシュタイン距離の閾値
    UseDoubleCheck As Boolean       '複数の方式でクロスチェックするか
End Type

Public Function GetWordStem(ByVal word As String, Optional ByVal config As StemmingConfig) As String
    '*** 設定されたメソッドに基づいて語幹を抽出 ***
    Select Case config.Method
        Case StemmingMethod.Porter
            GetWordStem = PorterStemmer(word)
            
        Case StemmingMethod.Levenshtein
            GetWordStem = LevenshteinBasedStem(word)
            
        Case StemmingMethod.Hybrid
            GetWordStem = HybridStemmer(word)
            
        Case Else
            GetWordStem = HybridStemmer(word)  'デフォルト
    End Select
End Function

Private Function LevenshteinBasedStem(ByVal word As String, Optional ByVal threshold As Double = 0.8) As String
    '*** レーベンシュタイン距離に基づく語幹抽出 ***
    Dim stem As String
    stem = LCase(Trim(word))
    
    '*** 共通の接尾辞を確認 ***
    Dim suffixes() As String
    suffixes = Split(COMMON_SUFFIXES, ",")
    
    Dim bestStem As String
    bestStem = stem
    Dim maxSimilarity As Double
    maxSimilarity = 0
    
    Dim i As Long
    For i = 0 To UBound(suffixes)
        If Len(stem) > Len(suffixes(i)) + 2 Then
            If Right(stem, Len(suffixes(i))) = suffixes(i) Then
                Dim candidateStem As String
                candidateStem = Left(stem, Len(stem) - Len(suffixes(i)))
                
                '*** レーベンシュタイン距離で類似度を計算 ***
                Dim similarity As Double
                similarity = 1 - (LevenshteinDistance(candidateStem, stem) / Len(stem))
                
                If similarity > maxSimilarity And similarity >= threshold Then
                    maxSimilarity = similarity
                    bestStem = candidateStem
                End If
            End If
        End If
    Next i
    
    LevenshteinBasedStem = bestStem
End Function

Public Function CompareWords(ByVal word1 As String, ByVal word2 As String, _
                           Optional ByVal config As StemmingConfig) As Boolean
    '*** 設定に基づいて単語を比較 ***
    Select Case config.Method
        Case StemmingMethod.Porter
            CompareWords = (GetWordStem(word1, config) = GetWordStem(word2, config))
            
        Case StemmingMethod.Levenshtein
            Dim similarity As Double
            similarity = 1 - (LevenshteinDistance(word1, word2) / _
                            Application.Max(Len(word1), Len(word2)))
            CompareWords = (similarity >= config.LevenshteinThreshold)
            
        Case StemmingMethod.Hybrid
            If config.UseDoubleCheck Then
                '*** 両方の方式でチェック ***
                Dim porterMatch As Boolean
                porterMatch = (GetWordStem(word1, config) = GetWordStem(word2, config))
                
                Dim levenMatch As Boolean
                levenMatch = (1 - (LevenshteinDistance(word1, word2) / _
                                 Application.Max(Len(word1), Len(word2)))) >= _
                                 config.LevenshteinThreshold
                
                CompareWords = porterMatch Or levenMatch
            Else
                CompareWords = (GetWordStem(word1, config) = GetWordStem(word2, config))
            End If
            
        Case Else
            CompareWords = (GetWordStem(word1, config) = GetWordStem(word2, config))
    End Select
End Function

'*** 設定を初期化する関数 ***
Public Function InitializeStemmingConfig(Optional ByVal method As StemmingMethod = StemmingMethod.Hybrid, _
                                       Optional ByVal threshold As Double = 0.8, _
                                       Optional ByVal useDoubleCheck As Boolean = False) As StemmingConfig
    With InitializeStemmingConfig
        .Method = method
        .LevenshteinThreshold = threshold
        .UseDoubleCheck = useDoubleCheck
    End With
End Function

Private Sub RemoveUnrelatedWordsWithConfig(ByVal ws As Worksheet, _
                                         ByVal targetWord As String, _
                                         ByVal config As StemmingConfig)
    '*** 関連のない単語を削除 ***
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' まず、対象単語と同じ語幹を持つ単語を削除
    Dim targetStem As String
    targetStem = GetWordStem(targetWord, config)
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim currentWord As String
        currentWord = ws.Cells(i, "D").Value
        
        ' 対象単語と完全に同じ場合は削除
        If LCase(Trim(currentWord)) = LCase(Trim(targetWord)) Then
            ws.Rows(i).Delete
            GoTo NextIteration
        End If
        
        ' 語幹が同じ場合は削除
        Dim currentStem As String
        currentStem = GetWordStem(currentWord, config)
        
        If currentStem = targetStem Then
            ws.Rows(i).Delete
            GoTo NextIteration
        End If
        
        ' 結果単語間で語幹が同じものを削除
        Dim j As Long
        For j = i - 1 To 2 Step -1
            Dim compareWord As String
            compareWord = ws.Cells(j, "D").Value
            
            If compareWord <> "" Then
                Dim compareStem As String
                compareStem = GetWordStem(compareWord, config)
                
                If currentStem = compareStem Then
                    ws.Rows(i).Delete
                    GoTo NextIteration
                End If
            End If
        Next j
        
NextIteration:
    Next i
End Sub

Private Sub RemoveDerivativesWithConfig(ByVal ws As Worksheet, _
                                      ByVal targetStem As String, _
                                      ByVal config As StemmingConfig)
    '*** 派生語を削除 ***
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim currentWord As String
        currentWord = ws.Cells(i, "D").Value
        
        ' 現在の単語の語幹を取得
        Dim currentStem As String
        currentStem = GetWordStem(currentWord, config)
        
        ' 語幹が一致する場合は削除
        If currentStem = targetStem Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Private Sub CopyInitialResults(ByVal tmpSheet As Worksheet, _
                             ByVal sourceSheet As Worksheet, _
                             ByVal targetGrade As String)
    '*** 初期データをコピー ***
    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' ヘッダーのコピー
    sourceSheet.Range("A1:F1").Copy tmpSheet.Range("A1")
    
    ' データのフィルタリングとコピー
    Dim i As Long
    Dim writeRow As Long
    writeRow = 2
    
    For i = 2 To lastRow
        If sourceSheet.Cells(i, "A").Value = targetGrade Then
            sourceSheet.Range("A" & i & ":F" & i).Copy _
                tmpSheet.Range("A" & writeRow)
            writeRow = writeRow + 1
        End If
    Next i
End Sub

Private Sub ProcessResultSet(ByVal wsTarget As Worksheet, _
                           ByVal wsList As Worksheet, _
                           ByVal targetWord As String, _
                           ByVal targetGrade As String, _
                           ByVal config As StemmingConfig, _
                           ByVal startCol As String, _
                           ByVal methodName As String)
                           
    '*** 一時シートを作成して処理 ***
    Dim tmpSheet As Worksheet
    Set tmpSheet = ThisWorkbook.Worksheets.Add
    
    '*** 初期データのコピー ***
    CopyInitialResults tmpSheet, wsList, targetGrade
    
    '*** フィルタリング処理 ***
    Dim targetStem As String
    targetStem = GetWordStem(targetWord, config)
    
    RemoveUnrelatedWordsWithConfig tmpSheet, targetWord, config
    RemoveDerivativesWithConfig tmpSheet, targetStem, config
    
    '*** 結果を目的の列にコピー ***
    CopyResults tmpSheet, wsTarget, startCol
    
    '*** 一時シートの削除 ***
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True
    
    '*** 結果数を表示 ***
    DisplayResultCount wsTarget, startCol, methodName
End Sub

Private Sub CopyResults(ByVal sourceWs As Worksheet, _
                       ByVal targetWs As Worksheet, _
                       ByVal startCol As String)
    '*** 結果のコピー ***
    Dim lastRow As Long
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row
    
    If lastRow >= 2 Then
        Dim targetRange As Range
        Set targetRange = targetWs.Range(startCol & "6")
        sourceWs.Range("A2:F" & lastRow).Copy targetRange
    End If
End Sub

Private Sub DisplayResultCount(ByVal ws As Worksheet, _
                             ByVal startCol As String, _
                             ByVal methodName As String)
    '*** 結果数のカウントと表示 ***
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    
    Dim resultCount As Long
    resultCount = lastRow - 5
    
    ws.Range(startCol & "4").Value = methodName & ": " & resultCount & "件"
End Sub

Private Sub HighlightUniqueResults(ByVal ws As Worksheet)
    '*** ユニークな結果をハイライト ***
    Dim porterLastRow As Long, levenLastRow As Long, hybridLastRow As Long
    Dim i As Long, found As Boolean
    Dim word As String
    
    porterLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    levenLastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    hybridLastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    '*** Porter方式のユニークな結果をハイライト ***
    For i = 6 To porterLastRow
        word = LCase(Trim(ws.Cells(i, "D").Value))
        If word <> "" Then
            found = IsWordInRange(ws.Range("K6:K" & levenLastRow), word) Or _
                   IsWordInRange(ws.Range("R6:R" & hybridLastRow), word)
            If Not found Then
                ws.Range("A" & i & ":G" & i).Interior.Color = RGB(217, 241, 255)
            End If
        End If
    Next i
    
    '*** Levenshtein方式のユニークな結果をハイライト ***
    For i = 6 To levenLastRow
        word = LCase(Trim(ws.Cells(i, "K").Value))
        If word <> "" Then
            found = IsWordInRange(ws.Range("D6:D" & porterLastRow), word) Or _
                   IsWordInRange(ws.Range("R6:R" & hybridLastRow), word)
            If Not found Then
                ws.Range("H" & i & ":N" & i).Interior.Color = RGB(217, 241, 255)
            End If
        End If
    Next i
    
    '*** ハイブリッド方式のユニークな結果をハイライト ***
    For i = 6 To hybridLastRow
        word = LCase(Trim(ws.Cells(i, "R").Value))
        If word <> "" Then
            found = IsWordInRange(ws.Range("D6:D" & porterLastRow), word) Or _
                   IsWordInRange(ws.Range("K6:K" & levenLastRow), word)
            If Not found Then
                ws.Range("O" & i & ":U" & i).Interior.Color = RGB(217, 241, 255)
            End If
        End If
    Next i
End Sub

Private Function IsWordInRange(ByVal rng As Range, ByVal word As String) As Boolean
    '*** 指定された範囲に単語が存在するかチェック ***
    Dim cell As Range
    For Each cell In rng
        If LCase(Trim(cell.Value)) = word Then
            IsWordInRange = True
            Exit Function
        End If
    Next cell
    IsWordInRange = False
End Function

Private Function PorterStemmer(ByVal word As String) As String
    '*** Porter Stemmerアルゴリズムの実装 ***
    Dim stem As String
    stem = LCase(Trim(word))
    
    '*** Step 1a ***
    If Right(stem, 4) = "sses" Then
    If Right(stem, 4) = "sses" Then
        stem = Left(stem, Len(stem) - 2)
    ElseIf Right(stem, 3) = "ies" Then
        stem = Left(stem, Len(stem) - 2)
    ElseIf Right(stem, 2) = "ss" Then
        ' Do nothing
    ElseIf Right(stem, 1) = "s" Then
        stem = Left(stem, Len(stem) - 1)
    End If
    
    '*** Step 1b ***
    Dim hasSuffix As Boolean
    hasSuffix = False
    
    If Right(stem, 3) = "eed" Then
        If MeasureCount(Left(stem, Len(stem) - 3)) > 0 Then
            stem = Left(stem, Len(stem) - 1)
        End If
    ElseIf Right(stem, 2) = "ed" Then
        If ContainsVowel(Left(stem, Len(stem) - 2)) Then
            stem = Left(stem, Len(stem) - 2)
            hasSuffix = True
        End If
    ElseIf Right(stem, 3) = "ing" Then
        If ContainsVowel(Left(stem, Len(stem) - 3)) Then
            stem = Left(stem, Len(stem) - 3)
            hasSuffix = True
        End If
    End If
    
    If hasSuffix Then
        If Right(stem, 2) = "at" Or Right(stem, 2) = "bl" Or Right(stem, 2) = "iz" Then
            stem = stem & "e"
        ElseIf DoubleSuffix(stem) And Right(stem, 1) <> "l" And Right(stem, 1) <> "s" And Right(stem, 1) <> "z" Then
            stem = Left(stem, Len(stem) - 1)
        ElseIf MeasureCount(stem) = 1 And EndsCVC(stem) Then
            stem = stem & "e"
        End If
    End If
    
    '*** Step 1c ***
    If ContainsVowel(Left(stem, Len(stem) - 1)) And Right(stem, 1) = "y" Then
        stem = Left(stem, Len(stem) - 1) & "i"
    End If
    
    '*** Step 2 ***
    Select Case Right(stem, 7)
        Case "ational"
            If MeasureCount(Left(stem, Len(stem) - 7)) > 0 Then
                stem = Left(stem, Len(stem) - 7) & "ate"
            End If
        Case "tional"
            If MeasureCount(Left(stem, Len(stem) - 6)) > 0 Then
                stem = Left(stem, Len(stem) - 6) & "tion"
            End If
    End Select
    
    Select Case Right(stem, 6)
        Case "ization"
            If MeasureCount(Left(stem, Len(stem) - 6)) > 0 Then
                stem = Left(stem, Len(stem) - 6) & "ize"
            End If
    End Select
    
    Select Case Right(stem, 5)
        Case "ation"
            If MeasureCount(Left(stem, Len(stem) - 5)) > 0 Then
                stem = Left(stem, Len(stem) - 5) & "ate"
            End If
    End Select
    
    '*** Step 3 ***
    Select Case Right(stem, 5)
        Case "alize"
            If MeasureCount(Left(stem, Len(stem) - 5)) > 0 Then
                stem = Left(stem, Len(stem) - 5) & "al"
            End If
    End Select
    
    Select Case Right(stem, 4)
        Case "ator"
            If MeasureCount(Left(stem, Len(stem) - 4)) > 0 Then
                stem = Left(stem, Len(stem) - 4) & "ate"
            End If
    End Select
    
    '*** Step 4 ***
    If Right(stem, 2) = "ic" Then
        If MeasureCount(Left(stem, Len(stem) - 2)) > 1 Then
            stem = Left(stem, Len(stem) - 2)
        End If
    End If
    
    PorterStemmer = stem
End Function

Private Function HybridStemmer(ByVal word As String) As String
    '*** Porter StemmingとLevenshtein距離を組み合わせたハイブリッド方式 ***
    Dim porterStem As String
    porterStem = PorterStemmer(word)
    
    Dim levenStem As String
    levenStem = LevenshteinBasedStem(word, 0.8)
    
    '*** 二つの結果を比較して、より短い方を採用 ***
    If Len(porterStem) <= Len(levenStem) Then
        HybridStemmer = porterStem
    Else
        HybridStemmer = levenStem
    End If
End Function

Private Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    '*** レーベンシュタイン距離を計算 ***
    Dim i As Long, j As Long
    Dim m As Long, n As Long
    Dim cost As Long
    
    '*** 文字列の長さを取得 ***
    m = Len(s1)
    n = Len(s2)
    
    '*** 配列の初期化 ***
    Dim d() As Long
    ReDim d(m, n)
    
    '*** 初期値の設定 ***
    For i = 0 To m
        d(i, 0) = i
    Next i
    
    For j = 0 To n
        d(0, j) = j
    Next j
    
    '*** 距離の計算 ***
    For i = 1 To m
        For j = 1 To n
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            
            d(i, j) = Application.Min( _
                d(i - 1, j) + 1, _
                d(i, j - 1) + 1, _
                d(i - 1, j - 1) + cost)
        Next j
    Next i
    
    LevenshteinDistance = d(m, n)
End Function

Private Function MeasureCount(ByVal stem As String) As Long
    '*** 語幹の母音-子音のシーケンス数をカウント ***
    Dim count As Long
    count = 0
    Dim hasVowel As Boolean
    hasVowel = False
    Dim i As Long
    
    For i = 1 To Len(stem)
        If IsVowel(Mid(stem, i, 1)) Then
            hasVowel = True
        ElseIf hasVowel Then
            count = count + 1
            hasVowel = False
        End If
    Next i
    
    MeasureCount = count
End Function

Private Function ContainsVowel(ByVal stem As String) As Boolean
    '*** 文字列に母音が含まれているかチェック ***
    Dim i As Long
    For i = 1 To Len(stem)
        If IsVowel(Mid(stem, i, 1)) Then
            ContainsVowel = True
            Exit Function
        End If
    Next i
    ContainsVowel = False
End Function

Private Function IsVowel(ByVal c As String) As Boolean
    '*** 母音かどうかをチェック ***
    Select Case LCase(c)
        Case "a", "e", "i", "o", "u"
            IsVowel = True
        Case Else
            IsVowel = False
    End Select
End Function

Private Function DoubleSuffix(ByVal stem As String) As Boolean
    '*** 二重子音で終わっているかチェック ***
    If Len(stem) < 2 Then
        DoubleSuffix = False
        Exit Function
    End If
    
    Dim lastChar As String
    lastChar = Right(stem, 1)
    
    If lastChar = Right(stem, 2) \ 2 Then
        DoubleSuffix = True
    Else
        DoubleSuffix = False
    End If
End Function

Private Function EndsCVC(ByVal stem As String) As Boolean
    '*** 子音-母音-子音で終わっているかチェック ***
    If Len(stem) < 3 Then
        EndsCVC = False
        Exit Function
    End If
    
    Dim last3 As String
    last3 = Right(stem, 3)
    
    If Not IsVowel(Mid(last3, 1, 1)) And _
       IsVowel(Mid(last3, 2, 1)) And _
       Not IsVowel(Mid(last3, 3, 1)) And _
       Mid(last3, 3, 1) <> "w" And _
       Mid(last3, 3, 1) <> "x" And _
       Mid(last3, 3, 1) <> "y" Then
        EndsCVC = True
    Else
        EndsCVC = False
    End If
End Function

Public Sub SearchRelatedWordsMultiMethod()
    '*** 複数の手法で検索を実行 ***
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim targetWord As String
    Dim targetGrade As String
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsTarget = ThisWorkbook.Sheets(4)
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** 検索条件の取得とチェック ***
    targetWord = LCase(Trim(wsTarget.Range("D2").Value))
    targetGrade = Trim(wsTarget.Range("C2").Value)
    
    If targetWord = "" Or targetGrade = "" Then
        MsgBox "C2セルに級、D2セルに検索する単語を入力してください。", vbExclamation
        Exit Sub
    End If
    
    '*** 結果エリアの準備 ***
    PrepareMultiResultArea wsTarget
    
    '*** 各手法での検索実行 ***
    Dim porterConfig As StemmingConfig
    porterConfig = InitializeStemmingConfig(StemmingMethod.Porter)
    ProcessResultSet wsTarget, wsList, targetWord, targetGrade, porterConfig, "A", "Porter方式"
    
    Dim levenConfig As StemmingConfig
    levenConfig = InitializeStemmingConfig(StemmingMethod.Levenshtein, 0.7)
    ProcessResultSet wsTarget, wsList, targetWord, targetGrade, levenConfig, "H", "レーベンシュタイン方式"
    
    Dim hybridConfig As StemmingConfig
    hybridConfig = InitializeStemmingConfig(StemmingMethod.Hybrid, 0.8, True)
    ProcessResultSet wsTarget, wsList, targetWord, targetGrade, hybridConfig, "O", "ハイブリッド方式"
    
    '*** ユニークな結果のハイライト ***
    HighlightUniqueResults wsTarget
    
    '*** 結果数の表示 ***
    ShowResultCounts wsTarget
    
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Sub ShowResultCounts(ByVal ws As Worksheet)
    '*** 各方式の結果数を表示 ***
    Dim porterCount As Long, levenCount As Long, hybridCount As Long
    Dim porterUnique As Long, levenUnique As Long, hybridUnique As Long
    
    porterCount = CountResults(ws, "A")
    levenCount = CountResults(ws, "H")
    hybridCount = CountResults(ws, "O")
    
    porterUnique = CountUniqueResults(ws, "A")
    levenUnique = CountUniqueResults(ws, "H")
    hybridUnique = CountUniqueResults(ws, "O")
    
    With ws
        .Range("A3").Value = "Porter方式: " & porterCount & "件 (ユニーク: " & porterUnique & "件)"
        .Range("H3").Value = "レーベンシュタイン方式: " & levenCount & "件 (ユニーク: " & levenUnique & "件)"
        .Range("O3").Value = "ハイブリッド方式: " & hybridCount & "件 (ユニーク: " & hybridUnique & "件)"
    End With
End Sub

Private Function CountResults(ByVal ws As Worksheet, ByVal startCol As String) As Long
    '*** 結果数をカウント ***
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    CountResults = lastRow - 5
End Function

Private Function CountUniqueResults(ByVal ws As Worksheet, ByVal startCol As String) As Long
    '*** ユニークな結果数をカウント ***
    Dim lastRow As Long, i As Long, count As Long
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    count = 0
    
    For i = 6 To lastRow
        If ws.Range(startCol & i).Interior.Color = RGB(217, 241, 255) Then
            count = count + 1
        End If
    Next i
    
    CountUniqueResults = count
End Function