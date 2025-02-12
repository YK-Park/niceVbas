
'word target tool 250212-01
Option Explicit

'*** 類似度を計算するための定数 ***
Private Const SIMILARITY_THRESHOLD As Double = 0.7    '類似度の閾値

'*** 語幹の抽出用の配列定数 ***
Private Const COMMON_SUFFIXES As String = "ed,ing,s,es,er,est,ly,ment,ness,ful,less,able,ible,al,ial,ic,ical,ish,like,ive,ative,itive"

'*** 検索結果を一時保存するための型 ***
Private Type SearchResult
    Row As Long        '元データの行番号
    Word As String     '単語
    Stem As String     '語幹
    Length As Long     '単語の長さ
End Type

Public Sub SearchRelatedWords()
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim targetWord As String
    Dim targetStem As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    Dim candidates() As SearchResult
    Dim candidateCount As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsTarget = ThisWorkbook.Sheets("ターゲット候補")
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    targetWord = LCase(Trim(wsTarget.Range("D2").Value))
    If targetWord = "" Then
        MsgBox "検索する単語を入力してください。", vbExclamation
        Exit Sub
    End If
    
    targetStem = GetWordStem(targetWord)
    
    '*** 結果エリアのクリア ***
    wsTarget.Range("J:O").ClearContents
    
    lastRow = wsList.Cells(wsList.Rows.Count, "D").End(xlUp).Row
    
    '*** 候補を格納する配列の初期化 ***
    ReDim candidates(1 To lastRow) As SearchResult
    candidateCount = 0
    
    '*** 第一段階：候補の収集 ***
    For i = 2 To lastRow
        Dim currentWord As String
        currentWord = LCase(Trim(GetLongestWord(wsList.Cells(i, "D").Value)))
        
        If currentWord <> "" And Len(currentWord) >= Len(targetWord) Then
            '*** 同じ単語でなく、類似度が閾値未満の場合は候補とする ***
            If currentWord <> targetWord And _
               CalculateSimilarity(currentWord, targetWord) < SIMILARITY_THRESHOLD Then
                
                candidateCount = candidateCount + 1
                With candidates(candidateCount)
                    .Row = i
                    .Word = currentWord
                    .Stem = GetWordStem(currentWord)
                    .Length = Len(currentWord)
                End With
            End If
        End If
    Next i
    
    '*** 第二段階：重複する語幹の排除と結果の書き込み ***
    resultRow = 1
    If candidateCount > 0 Then
        '*** 語幹でグループ化し、各グループの代表を選択 ***
        Dim j As Long, k As Long
        Dim processed() As Boolean
        ReDim processed(1 To candidateCount) As Boolean
        
        For j = 1 To candidateCount
            If Not processed(j) Then
                '*** まだ処理していない候補の場合 ***
                '*** 同じ語幹を持つ他の候補をマーク ***
                For k = j + 1 To candidateCount
                    If candidates(j).Stem = candidates(k).Stem Then
                        processed(k) = True
                    End If
                Next k
                
                '*** 結果の書き込み ***
                WriteResult wsList, wsTarget, candidates(j).Row, resultRow
                resultRow = resultRow + 1
                processed(j) = True
            End If
        Next j
    End If
    
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

'*** フレーズから最長の単語を抽出する関数 ***
Private Function GetLongestWord(ByVal phrase As String) As String
    '*** 空文字列チェック ***
    If Trim(phrase) = "" Then
        GetLongestWord = ""
        Exit Function
    End If
    
    Dim words() As String
    Dim i As Long
    Dim longestWord As String
    Dim maxLength As Long
    
    words = Split(Trim(phrase), " ")
    
    '*** 分割後の配列チェック ***
    If UBound(words) < 0 Then
        GetLongestWord = Trim(phrase)
        Exit Function
    End If
    
    maxLength = 0
    longestWord = ""
    
    For i = 0 To UBound(words)
        If Len(Trim(words(i))) > maxLength Then
            maxLength = Len(Trim(words(i)))
            longestWord = Trim(words(i))
        End If
    Next i
    
    GetLongestWord = longestWord
End Function

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