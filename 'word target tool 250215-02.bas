'word target tool 250215-02

Const MAX_RESULTS As Long = 100
Public Sub SearchRelatedWords()
    '*** メイン処理を行うサブルーチン ***
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    '*** ワークシートの設定 ***
    Dim wsTarget As Worksheet
    Dim wsList As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(4)
    Set wsList = ThisWorkbook.Sheets("単語リスト")
    
    '*** データの読み込み ***
    Dim lastSearchRow As Long, lastListRow As Long
    lastSearchRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    lastListRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    
    '*** 検索データと単語リストを配列に読み込む ***
    Dim searchWords() As Variant
    Dim listData() As Variant
    searchWords = wsTarget.Range("A2:A" & lastSearchRow).Value
    listData = wsList.Range("A2:F" & lastListRow).Value
    
    '*** 結果エリアのクリア ***
    wsTarget.Range("C1:XFD" & lastSearchRow).ClearContents
    
    '*** 検索結果の最大数を定数として定義 ***


'*** 結果を格納する配列の準備 ***
Dim results() As Variant
ReDim results(1 To lastSearchRow - 1, 1 To MAX_RESULTS) 

'그리고 루프 내에서:
If resultCount >= MAX_RESULTS Then Exit For  ' 최대 제한 체크
    '*** stemming結果をキャッシュするためのDictionary ***
    Dim stemCache As Object
    Set stemCache = CreateObject("Scripting.Dictionary")
    
    '*** 処理開始 ***
    Dim i As Long, j As Long, resultCount As Long
    
    For i = 1 To UBound(searchWords)
        If searchWords(i, 1) <> "" Then
            Application.StatusBar = "処理中... " & Format(i / UBound(searchWords) * 100, "0.0") & "%"
            DoEvents
            
            resultCount = 0
            Dim targetWord As String
            targetWord = LCase(Trim(CStr(searchWords(i, 1))))
            
            '*** キャッシュにない場合はstemming実行 ***
            If Not stemCache.Exists(targetWord) Then
                stemCache.Add targetWord, PorterStemmer(targetWord)
            End If
            
            For j = 1 To UBound(listData)
                If resultCount >= 100 Then Exit For  ' 最大100個まで
                
                Dim currentWord As String
                currentWord = LCase(Trim(CStr(listData(j, 4))))  ' D列の単語
                
                If currentWord <> "" And currentWord <> targetWord Then
                    If Not stemCache.Exists(currentWord) Then
                        stemCache.Add currentWord, PorterStemmer(currentWord)
                    End If
                    
                    If CompareWordsWithCache(targetWord, currentWord, stemCache) Then
                        resultCount = resultCount + 1
                        
                        '*** 結果を配列に格納 ***
                        Dim baseCol As Long
                        baseCol = (resultCount - 1) * 6
                        results(i, baseCol + 1) = listData(j, 1)  ' A列
                        results(i, baseCol + 2) = listData(j, 2)  ' B列
                        results(i, baseCol + 3) = listData(j, 3)  ' C列
                        results(i, baseCol + 4) = listData(j, 4)  ' D列
                        results(i, baseCol + 5) = listData(j, 5)  ' E列
                        results(i, baseCol + 6) = listData(j, 6)  ' F列
                    End If
                End If
            Next j
        End If
    Next i
    
    '*** 結果の書き込み ***
    If resultCount > 0 Then
        wsTarget.Range("C2").Resize(UBound(results), 100).Value = results
    End If
    
ExitSub:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Private Function CompareWordsWithCache(ByVal baseWord As String, ByVal compareWord As String, ByVal stemCache As Object) As Boolean
    '*** キャッシュされたstemming結果を使用して単語を比較 ***
    
    '*** 完全一致チェック ***
    If baseWord = compareWord Then
        CompareWordsWithCache = False
        Exit Function
    End If
    
    '*** ステミング結果の取得 ***
    Dim baseStem As String, compareStem As String
    baseStem = stemCache(baseWord)
    compareStem = stemCache(compareWord)
    
    '*** プレフィックスを除去 ***
    Dim baseWithoutPrefix As String, compareWithoutPrefix As String
    baseWithoutPrefix = RemovePrefix(baseWord)
    compareWithoutPrefix = RemovePrefix(compareWord)
    
    '*** 同じ語幹を持つ場合 ***
    If baseStem = compareStem Then
        '*** プレフィックスが異なる場合のみ含める ***
        If baseWithoutPrefix <> "" And compareWithoutPrefix <> "" Then
            If Not stemCache.Exists(baseWithoutPrefix) Then
                stemCache.Add baseWithoutPrefix, PorterStemmer(baseWithoutPrefix)
            End If
            If Not stemCache.Exists(compareWithoutPrefix) Then
                stemCache.Add compareWithoutPrefix, PorterStemmer(compareWithoutPrefix)
            End If
            
            If stemCache(baseWithoutPrefix) = stemCache(compareWithoutPrefix) Then
                CompareWordsWithCache = True
                Exit Function
            End If
        End If
        CompareWordsWithCache = False
        Exit Function
    End If
    
    CompareWordsWithCache = True
End Function