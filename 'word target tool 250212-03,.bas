'word target tool 250212-03
'급지정도 필수로 하여, 같은 급 내에서만 검색할 수 있도록 함
Public Sub SearchRelatedWords()
    '*** メイン処理を行うサブルーチン ***
    '*** Sheets(4)のA2に入力された単語を基に類似単語を検索し、C列以降に結果を表示する ***
    
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
    '*** 検索条件のチェック ***
    targetWord = LCase(Trim(wsTarget.Range("D2").Value))
    Dim targetGrade As String
    targetGrade = Trim(wsTarget.Range("C2").Value)

    If targetWord = "" Or targetGrade = "" Then
        MsgBox "C2セルに級、D2セルに検索する単語を入力してください。", vbExclamation
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
    Dim currentGrade As String
    
    currentWord = LCase(Trim(wsList.Cells(i, "D").Value))
    currentGrade = Trim(wsList.Cells(i, "C").Value)
    
    If currentWord <> "" And currentWord <> targetWord And currentGrade = targetGrade Then
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
            .AutoFilter Field:=3, '級 (C列)
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
    '*** 基準となる単語(baseWord)より短い単語は全て含める ***
    '*** 基準となる単語の派生語は除外 ***
    '*** 似ている単語（studio/study など）は含める ***
    
    '*** 両方の単語を小文字に変換して空白を削除 ***
    baseWord = LCase(Trim(baseWord))
    compareWord = LCase(Trim(compareWord))
    
    '*** 比較する単語が基準語より短い場合は常にTrue ***
    If Len(compareWord) < Len(baseWord) Then
        CompareWords = True
        Exit Function
    End If
    
    '*** 基準語が比較語の中に完全に含まれている場合は派生語として除外 ***
    If InStr(compareWord, baseWord) > 0 Then
        CompareWords = False
        Exit Function
    End If
    
    '*** それ以外の単語は含める ***
    CompareWords = True
End Function


'*** 結果から派生語を除去する関数 ***
Private Sub FilterDerivatives(ByVal ws As Worksheet, ByVal startRow As Long)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim word1 As String, word2 As String
    
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    
    For i = startRow To lastRow
        If Not processed(i) Then
            word1 = LCase(Trim(ws.Cells(i, "F").Value))
            
            For j = i + 1 To lastRow
                word2 = LCase(Trim(ws.Cells(j, "F").Value))
                
                '*** 派生語関係をチェック ***
                If InStr(word2, word1) > 0 Or InStr(word1, word2) > 0 Then
                    '*** 短い方を残し、長い方を削除 ***
                    If Len(word1) > Len(word2) Then
                        ws.Rows(i).Delete
                        Exit For
                    Else
                        ws.Rows(j).Delete
                        lastRow = lastRow - 1
                        j = j - 1
                    End If
                End If
            Next j
        End If
    Next i
End Sub
