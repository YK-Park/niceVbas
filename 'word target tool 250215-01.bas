'word target tool 250215-01'
' SearchRelatedWords (메인 서브루틴)
' CompareWords (단어 비교 함수)
' IsSameWordOrStem (동일 단어/어근 체크 함수)
' RemovePrefix (접두사 제거 함수)
' PorterStemmer (어근 추출 함수)
' 기타 보조 함수들 (ContainsVowel, IsDoubleConsonant, MeasureVC)

Public Sub SearchRelatedWords()
    '*** メイン処理を行うサブルーチン ***
    '*** A列の単語リストを基に類似単語を検索し、結果をC列から表示する ***
    
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Dim lastSearchRow As Long
    Dim lastListRow As Long
    Dim i As Long, j As Long
    Dim resultCol As Long
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    Set wsTarget = ThisWorkbook.Sheets(4)
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** 検索単語の最終行を取得 ***
    lastSearchRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    lastListRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    
    '*** 結果エリアのクリア（C列以降） ***
    wsTarget.Range("C:XFD").ClearContents
    
    '*** 各検索単語に対して処理 ***
    For i = 2 To lastSearchRow
        Dim targetWord As String
        targetWord = LCase(Trim(wsTarget.Cells(i, "A").Value))
        
        If targetWord <> "" Then
            Dim foundWords As Collection
            Set foundWords = New Collection
            
            '*** 類似単語の検索 ***
            For j = 2 To lastListRow
                Dim currentWord As String
                currentWord = LCase(Trim(wsList.Cells(j, "D").Value))
                
                If currentWord <> "" And currentWord <> targetWord Then
                    '*** A列の他の単語とのチェック ***
                    Dim isValidWord As Boolean
                    isValidWord = True
                    
                    '*** A列の他の単語との比較 ***
                    Dim k As Long
                    For k = 2 To lastSearchRow
                        If k <> i Then  '自分自身は除外
                            Dim otherWord As String
                            otherWord = LCase(Trim(wsTarget.Cells(k, "A").Value))
                            If otherWord <> "" Then
                                If IsSameWordOrStem(currentWord, otherWord) Then
                                    isValidWord = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                    
                    '*** 結果リストとの重複チェックと追加 ***
                    If isValidWord Then
                        If CompareWords(targetWord, currentWord) Then
                            Dim isDuplicate As Boolean
                            isDuplicate = False
                            
                            On Error Resume Next
                            foundWords.Add currentWord, currentWord
                            If Err.Number = 0 Then
                                '*** 結果の書き込み ***
                                resultCol = foundWords.Count
                                wsTarget.Cells(i, "C").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "A").Value  '級番号
                                wsTarget.Cells(i, "D").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "B").Value  'ユニーク番号
                                wsTarget.Cells(i, "E").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "C").Value  '級
                                wsTarget.Cells(i, "F").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "D").Value  '単語
                                wsTarget.Cells(i, "G").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "E").Value  '品詞
                                wsTarget.Cells(i, "H").Offset(0, (resultCol - 1) * 6).Value = wsList.Cells(j, "F").Value  '出題区分
                            End If
                            On Error GoTo ErrorHandler
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Function CompareWords(ByVal baseWord As String, ByVal compareWord As String) As Boolean
    '*** 単語比較関数 ***
    baseWord = LCase(Trim(baseWord))
    compareWord = LCase(Trim(compareWord))
    
    '*** 完全一致チェック ***
    If baseWord = compareWord Then
        CompareWords = False
        Exit Function
    End If
    
    '*** ステミング処理 ***
    Dim baseStem As String
    Dim compareStem As String
    baseStem = PorterStemmer(baseWord)
    compareStem = PorterStemmer(compareWord)
    
    '*** プレフィックスを除去 ***
    Dim baseWithoutPrefix As String
    Dim compareWithoutPrefix As String
    baseWithoutPrefix = RemovePrefix(baseWord)
    compareWithoutPrefix = RemovePrefix(compareWord)
    
    If baseStem = compareStem Then
        '*** プレフィックスのみが異なる関連単語の場合 ***
        If baseWithoutPrefix <> "" And compareWithoutPrefix <> "" Then
            If PorterStemmer(baseWithoutPrefix) = PorterStemmer(compareWithoutPrefix) Then
                CompareWords = True
                Exit Function
            End If
        End If
        CompareWords = False
        Exit Function
    End If
    
    '*** 異なる語幹を持つ単語は含める ***
    CompareWords = True
End Function

Private Function IsSameWordOrStem(ByVal word1 As String, ByVal word2 As String) As Boolean
    '*** 同じ単語または語幹かをチェック ***
    word1 = LCase(Trim(word1))
    word2 = LCase(Trim(word2))
    
    If word1 = word2 Then
        IsSameWordOrStem = True
        Exit Function
    End If
    
    If PorterStemmer(word1) = PorterStemmer(word2) Then
        IsSameWordOrStem = True
        Exit Function
    End If
    
    IsSameWordOrStem = False
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
End Function