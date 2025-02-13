'word target tool 250213-03 
'포터 스테밍 알고리즘 사용
Public Sub SearchRelatedWords()
    '*** メイン処理を行うサブルーチン ***
    '*** Sheets(4)のD2に入力された単語を基に類似単語を検索し、結果を表示する ***
    
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim targetWord As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsTarget = ThisWorkbook.Sheets(4)
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** 基準単語の取得 ***
    targetWord = LCase(Trim(wsTarget.Range("D2").Value))
    If targetWord = "" Then
        MsgBox "D2セルに検索する単語を入力してください。", vbExclamation
        Exit Sub
    End If
    
    '*** 結果エリアのクリア（ヘッダー行を除く） ***
    If wsTarget.FilterMode Then
        wsTarget.ShowAllData
    End If

    '*** データのクリア ***
    wsTarget.Range("A6:F" & wsTarget.Rows.Count).ClearContents
    
    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    resultRow = 2  '*** ヘッダー行の次から開始 ***
    
    '*** 第一段階：類似単語の検索 ***
    For i = 2 To lastRow
        Dim currentWord As String
        currentWord = LCase(Trim(wsList.Cells(i, "D").Value))
        
        If currentWord <> "" And currentWord <> targetWord Then
            If CompareWords(targetWord, currentWord) Then
                '*** 結果の書き込み ***
                wsTarget.Cells(resultRow + 4, "A").Value = wsList.Cells(i, "A").Value  '級番号
                wsTarget.Cells(resultRow + 4, "B").Value = wsList.Cells(i, "B").Value  'ユニーク番号
                wsTarget.Cells(resultRow + 4, "C").Value = wsList.Cells(i, "C").Value  '級
                wsTarget.Cells(resultRow + 4, "D").Value = wsList.Cells(i, "D").Value  '単語
                wsTarget.Cells(resultRow + 4, "E").Value = wsList.Cells(i, "E").Value  '品詞
                wsTarget.Cells(resultRow + 4, "F").Value = wsList.Cells(i, "F").Value  '出題区分
                resultRow = resultRow + 1
            End If
        End If
    Next i
    
    '*** 第二段階：派生語の除去 ***
    If resultRow > 2 Then
        FilterDerivatives wsTarget, 6  '*** ヘッダー行を除いて処理 ***
        With wsTarget.Range("A6:F" & resultRow + 4)
            .AutoFilter
            .AutoFilter Field:=5  '品詞 (E列)
        End With
    End If
    
ExitSub:
    Application.ScreenUpdating = True
    
    '*** 検索結果の報告 ***
    If resultRow > 2 Then
        MsgBox (resultRow - 2) & "件の類似単語が見つかりました。", vbInformation
    Else
        MsgBox "該当する単語は見つかりませんでした。", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Function CompareWords(ByVal baseWord As String, ByVal compareWord As String) As Boolean
    '*** 単語比較関数 ***
    '*** ポーターステミングを使用して語幹を比較 ***
    
    '*** 両方の単語を小文字に変換して空白を削除 ***
    baseWord = LCase(Trim(baseWord))
    compareWord = LCase(Trim(compareWord))
    
    '*** ステミング処理 ***
    Dim baseStem As String
    Dim compareStem As String
    
    baseStem = PorterStemmer(baseWord)
    compareStem = PorterStemmer(compareWord)
    
    '*** 語幹が同じ場合は除外 ***
    If baseStem = compareStem Then
        CompareWords = False
        Exit Function
    End If
    
    '*** プレフィックスチェック ***
    If HasCommonPrefix(baseWord, compareWord) Then
        CompareWords = False
        Exit Function
    End If
    
    CompareWords = True
End Function

Private Function HasCommonPrefix(ByVal word1 As String, ByVal word2 As String) As Boolean
    '*** 共通のプレフィックスをチェック ***
    Dim stem1 As String, stem2 As String
    
    '*** プレフィックスの抽出と語幹の比較 ***
    stem1 = RemovePrefix(word1)
    stem2 = RemovePrefix(word2)
    
    If stem1 <> "" And stem2 <> "" Then
        If PorterStemmer(stem1) = PorterStemmer(stem2) Then
            HasCommonPrefix = True
            Exit Function
        End If
    End If
    
    HasCommonPrefix = False
End Function

Private Function RemovePrefix(ByVal word As String) As String
    '*** プレフィックスを除去して語幹を返す ***
    Select Case True
        Case Left(word, 4) = "over"
            RemovePrefix = Mid(word, 5)
        Case Left(word, 5) = "under"
            RemovePrefix = Mid(word, 6)
        Case Left(word, 5) = "super"
            RemovePrefix = Mid(word, 6)
        Case Left(word, 3) = "pre"
            RemovePrefix = Mid(word, 4)
        Case Left(word, 2) = "re"
            RemovePrefix = Mid(word, 3)
        Case Left(word, 2) = "un"
            RemovePrefix = Mid(word, 3)
        Case Left(word, 2) = "in"
            RemovePrefix = Mid(word, 3)
        Case Left(word, 3) = "dis"
            RemovePrefix = Mid(word, 4)
        Case Left(word, 3) = "mis"
            RemovePrefix = Mid(word, 4)
        Case Left(word, 3) = "sub"
            RemovePrefix = Mid(word, 4)
        Case Else
            RemovePrefix = ""
    End Select
End Function

Private Sub FilterDerivatives(ByVal ws As Worksheet, ByVal startRow As Long)
    '*** 派生語フィルター（ポーターステミング使用） ***
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim word1 As String, word2 As String
    Dim stem1 As String, stem2 As String
    
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    For i = startRow To lastRow
        word1 = LCase(Trim(ws.Cells(i, "D").Value))
        If word1 <> "" Then
            stem1 = PorterStemmer(word1)
            
            For j = i + 1 To lastRow
                word2 = LCase(Trim(ws.Cells(j, "D").Value))
                If word2 <> "" Then
                    stem2 = PorterStemmer(word2)
                    
                    '*** 語幹が同じ場合は片方を削除 ***
                    If stem1 = stem2 Then
                        ws.Rows(j).Delete
                        lastRow = lastRow - 1
                        j = j - 1
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Private Function PorterStemmer(ByVal word As String) As String
    '*** ポーターステミングアルゴリズムの実装 ***
    Dim stem As String
    stem = LCase(Trim(word))
    
    '*** ステップ1a: 複数形と過去分詞の処理 ***
    If stem Like "*sses" Then
        stem = Left(stem, Len(stem) - 2)
    ElseIf stem Like "*ies" Then
        stem = Left(stem, Len(stem) - 2)
    ElseIf stem Like "*ss" Then
        '何もしない
    ElseIf stem Like "*s" Then
        stem = Left(stem, Len(stem) - 1)
    End If
    
    '*** ステップ1b: 過去形と進行形の処理 ***
    Dim m As Long
    If stem Like "*eed" Then
        m = MeasureVC(Left(stem, Len(stem) - 3))
        If m > 0 Then stem = Left(stem, Len(stem) - 1)
    ElseIf stem Like "*ed" Then
        If ContainsVowel(Left(stem, Len(stem) - 2)) Then
            stem = Left(stem, Len(stem) - 2)
            '二重子音の処理
            If stem Like "*at" Or stem Like "*bl" Or stem Like "*iz" Then
                stem = stem & "e"
            ElseIf IsDoubleConsonant(stem) And Not (Right(stem, 1) = "l" Or Right(stem, 1) = "s" Or Right(stem, 1) = "z") Then
                stem = Left(stem, Len(stem) - 1)
            End If
        End If
    ElseIf stem Like "*ing" Then
        If ContainsVowel(Left(stem, Len(stem) - 3)) Then
            stem = Left(stem, Len(stem) - 3)
            '二重子音の処理
            If stem Like "*at" Or stem Like "*bl" Or stem Like "*iz" Then
                stem = stem & "e"
            ElseIf IsDoubleConsonant(stem) And Not (Right(stem, 1) = "l" Or Right(stem, 1) = "s" Or Right(stem, 1) = "z") Then
                stem = Left(stem, Len(stem) - 1)
            End If
        End If
    End If
    
    '*** ステップ1c: y→iの変換 ***
    If stem Like "*y" And ContainsVowel(Left(stem, Len(stem) - 1)) Then
        stem = Left(stem, Len(stem) - 1) & "i"
    End If
    
    '*** ステップ2: 接尾辞の変換 ***
    m = MeasureVC(stem)
    If m > 0 Then
        Select Case True
            Case stem Like "*ational": stem = Left(stem, Len(stem) - 7) & "ate"
            Case stem Like "*tional": stem = Left(stem, Len(stem) - 6) & "tion"
            Case stem Like "*enci": stem = Left(stem, Len(stem) - 4) & "ence"
            Case stem Like "*anci": stem = Left(stem, Len(stem) - 4) & "ance"
            Case stem Like "*izer": stem = Left(stem, Len(stem) - 4) & "ize"
            Case stem Like "*abli": stem = Left(stem, Len(stem) - 4) & "able"
            Case stem Like "*alli": stem = Left(stem, Len(stem) - 4) & "al"
            Case stem Like "*entli": stem = Left(stem, Len(stem) - 5) & "ent"
            Case stem Like "*eli": stem = Left(stem, Len(stem) - 3) & "e"
            Case stem Like "*ousli": stem = Left(stem, Len(stem) - 5) & "ous"
        End Select
    End If
    
    '*** ステップ3: 更なる接尾辞の処理 ***
    m = MeasureVC(stem)
    If m > 0 Then
        Select Case True
            Case stem Like "*icate": stem = Left(stem, Len(stem) - 5) & "ic"
            Case stem Like "*ative": stem = Left(stem, Len(stem) - 5)
            Case stem Like "*alize": stem = Left(stem, Len(stem) - 5) & "al"
            Case stem Like "*iciti": stem = Left(stem, Len(stem) - 5) & "ic"
            Case stem Like "*ical": stem = Left(stem, Len(stem) - 4) & "ic"
            Case stem Like "*ful": stem = Left(stem, Len(stem) - 3)
            Case stem Like "*ness": stem = Left(stem, Len(stem) - 4)
        End Select
    End If
    
    PorterStemmer = stem
End Function

Private Function ContainsVowel(ByVal word As String) As Boolean
    '*** 母音を含むかどうかをチェック ***
    ContainsVowel = word Like "*a*" Or word Like "*e*" Or word Like "*i*" Or word Like "*o*" Or word Like "*u*"
End Function

Private Function IsDoubleConsonant(ByVal word As String) As Boolean
    '*** 二重子音かどうかをチェック ***
    If Len(word) < 2 Then Exit Function
    Dim lastChar As String, secondLastChar As String
    lastChar = Right(word, 1)
    secondLastChar = Mid(word, Len(word) - 1, 1)
    IsDoubleConsonant = (lastChar = secondLastChar) And Not ContainsVowel(lastChar)
End Function

Private Function MeasureVC(ByVal word As String) As Long
    '*** 母音-子音のシーケンス数を計測 ***
    Dim count As Long
    Dim i As Long
    Dim inVowel As Boolean
    
    count = 0
    inVowel = False
    
    For i = 1 To Len(word)
        If IsVowel(Mid(word, i, 1)) Then
            If Not inVowel Then
                inVowel = True
            End If
        Else
            If inVowel Then
                count = count + 1
                inVowel = False
            End If
        End If
        
    Next i
    
    MeasureVC = count
End Function

Private Function IsVowel(ByVal char As String) As Boolean
    '*** 母音かどうかをチェック ***
    IsVowel = char = "a" Or char = "e" Or char = "i" Or char = "o" Or char = "u"
End Function'*** メイン処理を行うサブルーチン ***
   